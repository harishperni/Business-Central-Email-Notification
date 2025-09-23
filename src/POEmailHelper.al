codeunit 50142 "PO Email Helper"
{
    var
        PostedWrReportId: Integer; // set where used (Arrived)

    // =========================
    // Public notification APIs
    // =========================

    procedure Notify_POCreated_OnRelease(PurchHeader: Record "Purchase Header")
    var
        subj: Text;
        body: TextBuilder;
        html: Text;
    begin
        subj := StrSubstNo('PO %1 Created', PurchHeader."No.");

        body.AppendLine('<p>Dear Colleague,</p>');
        body.AppendLine(StrSubstNo(
            '<p>PO <strong>#%1</strong> has been created per your request. You will receive another notification when your item(s) ship. The items and quantities are:</p>',
            Html(PurchHeader."No.")));
        body.AppendLine(BuildHtmlTableFromPOLines(PurchHeader));
        body.AppendLine(BuildFooter());

        html := WrapHtml(subj, body.ToText());
        SendEmail(subj, html, BuildRecipientListFromPO(PurchHeader));
    end;

    procedure Notify_Shipped_OnWhseReceiptCreated(WhseRcptHeader: Record "Warehouse Receipt Header"; RelatedPO: Record "Purchase Header")
    var
        subj: Text;
        body: TextBuilder;
        html: Text;
        eta: Text;
    begin
        if RelatedPO."Expected Receipt Date" <> 0D then
            eta := Format(RelatedPO."Expected Receipt Date");

        subj := StrSubstNo('PO %1 Shipped', RelatedPO."No.");

        body.AppendLine('<p>Dear Colleague,</p>');
        if eta <> '' then
            body.AppendLine(StrSubstNo(
                '<p>PO <strong>#%1</strong> has <strong>Shipped!!</strong> The estimated arrival date is <strong>%2</strong>. The items and quantities are:</p>',
                Html(RelatedPO."No."), Html(eta)))
        else
            body.AppendLine(StrSubstNo(
                '<p>PO <strong>#%1</strong> has <strong>Shipped!!</strong>. The items and quantities are:</p>',
                Html(RelatedPO."No.")));

        // Only WR lines belonging to this PO
        body.AppendLine(BuildHtmlTableFromWhseRcptLines(WhseRcptHeader, RelatedPO."No."));
        body.AppendLine(BuildFooter());

        html := WrapHtml(subj, body.ToText());
        SendEmail(subj, html, BuildRecipientListFromPO(RelatedPO));
    end;

    procedure Notify_Arrived_OnWhseReceiptPosted(PostedHdr: Record "Posted Whse. Receipt Header"; RelatedPO: Record "Purchase Header")
    var
        subj: Text;
        body: TextBuilder;
        html: Text;
        Email: Codeunit Email;
        EmailMsg: Codeunit "Email Message";
        ToList: List of [Text];
        TempBlob: Codeunit "Temp Blob";
        OutS: OutStream;
        InS: InStream;
        RecRef: RecordRef;
        FileName: Text[250];
        PostedLine: Record "Posted Whse. Receipt Line";
        LineCount: Integer;
    begin
        // Assign report id where used
        PostedWrReportId := 7308; // "Warehouse Posted Receipt"

        // ── LOG: start
        Session.LogMessage(
            'arrived.start',
            StrSubstNo('Arrived email start. PostedWR=%1, PO=%2', PostedHdr."No.", RelatedPO."No."),
            Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
            'PostedWR', PostedHdr."No.");

        subj := StrSubstNo('PO %1 Arrived', RelatedPO."No.");

        // Count posted lines for this posted receipt and this PO
        PostedLine.SetRange("No.", PostedHdr."No.");
        PostedLine.SetRange("Source Type", Database::"Purchase Line");
        PostedLine.SetRange("Source No.", RelatedPO."No.");
        if PostedLine.FindSet() then
            repeat
                if PostedLine."Item No." <> '' then
                    LineCount += 1;
            until PostedLine.Next() = 0;

        // ── LOG: line count
        Session.LogMessage(
            'arrived.count',
            StrSubstNo('Arrived posted lines: %1 (PostedWR=%2, PO=%3)', LineCount, PostedHdr."No.", RelatedPO."No."),
            Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
            'Count', Format(LineCount));

        body.AppendLine('<p>Dear Colleague,</p>');
        body.AppendLine(StrSubstNo(
            '<p>PO <strong>#%1</strong> has <strong>Arrived</strong> at our Bestway USA Chandler Warehouse!! ' +
            'Please allow 3–5 days for Container Unloading &amp; Putaways. The posted items and quantities are:</p>',
            Html(RelatedPO."No.")));

        // HTML table of ALL posted lines for this posted receipt (filtered to this PO)
        body.AppendLine(BuildHtmlTableFromPostedWRLines(PostedHdr, RelatedPO."No."));

        body.AppendLine('<p>Attached is the Posted Warehouse Receipt PDF for your reference.</p>');
        body.AppendLine(BuildFooter());

        html := WrapHtml(subj, body.ToText());

        // Recipients
        ToList := BuildRecipientListFromPO(RelatedPO);
        if ToList.Count() = 0 then begin
            Session.LogMessage(
                'arrived.norecip',
                StrSubstNo('No recipients resolved for PO %1', RelatedPO."No."),
                Verbosity::Warning, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
                'PO', RelatedPO."No.");
            exit;
        end;

        // Create email with HTML body
        EmailMsg.Create(ToList, subj, html, true);

        // Generate and attach PDF
        RecRef.GetTable(PostedHdr);
        TempBlob.CreateOutStream(OutS);
        Report.SaveAs(PostedWrReportId, '', ReportFormat::Pdf, OutS, RecRef);
        TempBlob.CreateInStream(InS);

        FileName := StrSubstNo('PostedWhseReceipt_%1.pdf', PostedHdr."No.");
        EmailMsg.AddAttachment(FileName, 'application/pdf', InS);

        // ── LOG: pdf attached
        Session.LogMessage(
            'arrived.pdf.ok',
            StrSubstNo('PDF attached. ReportId=%1, PostedWR=%2', Format(PostedWrReportId), PostedHdr."No."),
            Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
            'ReportId', Format(PostedWrReportId));

        Email.Send(EmailMsg, Enum::"Email Scenario"::Default);

        // ── LOG: sent
        Session.LogMessage(
            'arrived.sent',
            StrSubstNo('Arrived email sent. PostedWR=%1, PO=%2, Lines=%3', PostedHdr."No.", RelatedPO."No.", LineCount),
            Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
            'PostedWR', PostedHdr."No.");
    end;

    // =========================
    // Rendering / utilities
    // =========================

    local procedure BuildHtmlTableFromPOLines(PurchHeader: Record "Purchase Header"): Text
    var
        PurchLine: Record "Purchase Line";
        b: TextBuilder;
    begin
        b.AppendLine('<table border="0" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%;max-width:900px;">');
        b.AppendLine('<tr style="background:#0a66c2;color:#fff;">' +
                     '<th align="left">Item No.</th>' +
                     '<th align="left">Description</th>' +
                     '<th align="right">Qty</th>' +
                     '<th align="left">Location</th>' +
                     '</tr>');

        PurchLine.SetRange("Document Type", PurchHeader."Document Type");
        PurchLine.SetRange("Document No.", PurchHeader."No.");
        if PurchLine.FindSet() then
            repeat
                if (PurchLine.Type = PurchLine.Type::Item) and (PurchLine."No." <> '') then
                    b.AppendLine(StrSubstNo(
                        '<tr><td>%1</td><td>%2</td><td align="right">%3</td><td>%4</td></tr>',
                        Html(PurchLine."No."),
                        Html(PurchLine.Description),
                        Html(Format(PurchLine.Quantity)),
                        Html(PurchLine."Location Code")));
            until PurchLine.Next() = 0;

        b.AppendLine('</table>');
        exit(b.ToText());
    end;

    local procedure BuildHtmlTableFromWhseRcptLines(WhseRcptHeader: Record "Warehouse Receipt Header"; PONo: Code[20]): Text
    var
        WhseRcptLine: Record "Warehouse Receipt Line";
        b: TextBuilder;
    begin
        b.AppendLine('<table border="0" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%;max-width:900px;">');
        b.AppendLine('<tr style="background:#0a66c2;color:#fff;">' +
                     '<th align="left">Item No.</th>' +
                     '<th align="left">Description</th>' +
                     '<th align="right">Qty</th>' +
                     '<th align="left">Location</th>' +
                     '</tr>');

        WhseRcptLine.SetRange("No.", WhseRcptHeader."No.");
        WhseRcptLine.SetRange("Source Type", Database::"Purchase Line");
        WhseRcptLine.SetRange("Source No.", PONo);
        if WhseRcptLine.FindSet() then
            repeat
                if WhseRcptLine."Item No." <> '' then
                    b.AppendLine(StrSubstNo(
                        '<tr><td>%1</td><td>%2</td><td align="right">%3</td><td>%4</td></tr>',
                        Html(WhseRcptLine."Item No."),
                        Html(WhseRcptLine.Description),
                        Html(Format(WhseRcptLine.Quantity)),
                        Html(WhseRcptLine."Location Code")));
            until WhseRcptLine.Next() = 0;

        b.AppendLine('</table>');
        exit(b.ToText());
    end;

    local procedure BuildHtmlTableFromPostedWRLines(PostedHdr: Record "Posted Whse. Receipt Header"; PONo: Code[20]): Text
    var
        PwrLine: Record "Posted Whse. Receipt Line";
        b: TextBuilder;
    begin
        b.AppendLine('<table border="0" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%;max-width:900px;">');
        b.AppendLine('<tr style="background:#0a66c2;color:#fff;">' +
                     '<th align="left">Item No.</th>' +
                     '<th align="left">Description</th>' +
                     '<th align="right">Qty</th>' +
                     '<th align="left">Location</th>' +
                     '</tr>');

        PwrLine.SetRange("No.", PostedHdr."No.");
        PwrLine.SetRange("Source Type", Database::"Purchase Line");
        PwrLine.SetRange("Source No.", PONo);

        if PwrLine.FindSet() then
            repeat
                if PwrLine."Item No." <> '' then
                    b.AppendLine(StrSubstNo(
                        '<tr><td>%1</td><td>%2</td><td align="right">%3</td><td>%4</td></tr>',
                        Html(PwrLine."Item No."),
                        Html(PwrLine.Description),
                        Html(Format(PwrLine.Quantity)),
                        Html(PwrLine."Location Code")));
            until PwrLine.Next() = 0;

        b.AppendLine('</table>');
        exit(b.ToText());
    end;

    // ---------------- Recipients ----------------

    local procedure BuildRecipientListFromPO(PurchHeader: Record "Purchase Header"): List of [Text]
    var
        list: List of [Text];
        addr: Text;
    begin
        addr := ResolvePOPrimaryEmail(PurchHeader);
        if addr <> '' then
            list.Add(addr);
        exit(list);
    end;

    local procedure ResolvePOPrimaryEmail(PurchHeader: Record "Purchase Header"): Text
    var
        Vendor: Record Vendor;
        Contact: Record Contact;
    begin
        if PurchHeader."Buy-from Contact No." <> '' then
            if Contact.Get(PurchHeader."Buy-from Contact No.") then
                if Contact."E-Mail" <> '' then
                    exit(Contact."E-Mail");

        if Vendor.Get(PurchHeader."Buy-from Vendor No.") then
            if Vendor."E-Mail" <> '' then
                exit(Vendor."E-Mail");

        exit('');
    end;

    // ---------------- Send email ----------------

    local procedure SendEmail(SubjectTxt: Text; BodyHtml: Text; ToList: List of [Text])
    var
        Email: Codeunit Email;
        EmailMessage: Codeunit "Email Message";
    begin
        if ToList.Count() = 0 then
            exit;

        EmailMessage.Create(ToList, SubjectTxt, BodyHtml, true);
        Email.Send(EmailMessage, Enum::"Email Scenario"::Default);
    end;

    // ---------------- HTML helpers ----------------

    local procedure WrapHtml(Title: Text; BodyHtml: Text): Text
    var
        b: TextBuilder;
    begin
        b.AppendLine('<!DOCTYPE html><html><head><meta charset="utf-8"/>');
        b.AppendLine('<title>' + Html(Title) + '</title>');
        b.AppendLine('<style>body{font-family:Segoe UI,Arial,Helvetica,sans-serif;font-size:13px;line-height:1.35}</style>');
        b.AppendLine('</head><body>');
        b.AppendLine(BodyHtml);
        b.AppendLine('</body></html>');
        exit(b.ToText());
    end;

    local procedure BuildFooter(): Text
    begin
        exit(
          '<p>If you have any questions, please contact your Supply Chain Team: ' +
          '<a href="mailto:cyndy.peterson@bestwaycorp.us">cyndy.peterson@bestwaycorp.us</a> ' +
          'and/or <a href="mailto:Eric.Eichstaedt@bestwaycorp.us">Eric.Eichstaedt@bestwaycorp.us</a>.</p>');
    end;

    // Minimal HTML encoding (no ConvertStr; avoid reserved word "With")
    local procedure Html(t: Text): Text
    begin
        t := ReplaceAll(t, '&', '&amp;');
        t := ReplaceAll(t, '<', '&lt;');
        t := ReplaceAll(t, '>', '&gt;');
        t := ReplaceAll(t, '"', '&quot;');
        exit(t);
    end;

    local procedure ReplaceAll(S: Text; Find: Text; ReplaceWith: Text): Text
    var
        R: Text;
        P: Integer;
        Lf: Integer;
    begin
        if (Find = '') or (S = '') then
            exit(S);
        R := S;
        Lf := StrLen(Find);
        P := StrPos(R, Find);
        while P > 0 do begin
            R := CopyStr(R, 1, P - 1) + ReplaceWith + CopyStr(R, P + Lf);
            P := StrPos(R, Find);
        end;
        exit(R);
    end;
}