codeunit 50142 "PO Email Helper"
{
    var
        PostedWrReportId: Integer; // set where used (Arrived)

    // =========================
    // Public notification APIs
    // =========================

    procedure Notify_POCreated_OnRelease(PurchHeader: Record "Purchase Header")
    var
        PL: Record "Purchase Line";
        b: TextBuilder;
        html: Text;
        subj: Text;
    begin
        b.AppendLine('<p>Dear Colleague,</p>');
        b.AppendLine(StrSubstNo('<p>PO <strong>#%1</strong> has been created per your request. You will receive another notification when your item(s) ship. The items and quantities are:</p>', Html(PurchHeader."No.")));

        // table (borders, spacing, right-aligned Qty)
        b.AppendLine('<table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;width:100%;max-width:900px;margin:12px 0;border:1px solid #ccc;">');
        b.AppendLine('<tr style="background:#0a66c2;color:#fff;">' +
                     '<th align="left"  style="border:1px solid #ccc;">Item No.</th>' +
                     '<th align="left"  style="border:1px solid #ccc;">Description</th>' +
                     '<th align="right" style="border:1px solid #ccc;">Qty</th>' +
                     '<th align="left"  style="border:1px solid #ccc;">Location</th>' +
                     '</tr>');

        PL.SetRange("Document Type", PurchHeader."Document Type");
        PL.SetRange("Document No.", PurchHeader."No.");
        if PL.FindSet() then
            repeat
                if (PL.Type = PL.Type::Item) and (PL."No." <> '') then
                    b.AppendLine(StrSubstNo(
                        '<tr>' +
                        '<td style="border:1px solid #ccc;">%1</td>' +
                        '<td style="border:1px solid #ccc;">%2</td>' +
                        '<td style="border:1px solid #ccc;text-align:right;">%3</td>' +
                        '<td style="border:1px solid #ccc;">%4</td>' +
                        '</tr>',
                        Html(PL."No."),
                        Html(PL.Description),
                        Html(Format(PL.Quantity)),
                        Html(PL."Location Code")));
            until PL.Next() = 0;

        b.AppendLine('</table>');
        b.AppendLine('<p>If you have any questions, please contact your Supply Chain Team: ' +
                     '<a href="mailto:cyndy.peterson@bestwaycorp.us">cyndy.peterson@bestwaycorp.us</a> and/or ' +
                     '<a href="mailto:Eric.Eichstaedt@bestwaycorp.us">Eric.Eichstaedt@bestwaycorp.us</a>.</p>');

        subj := StrSubstNo('PO %1 Created', PurchHeader."No.");
        html := WrapHtml(subj, b.ToText());
        SendEmail(subj, html, BuildRecipientListFromPO(PurchHeader));
    end;

    // === SHIPPED (WR CREATED) ===============================================
    procedure Notify_Shipped_OnWhseReceiptCreated(WhseRcptHeader: Record "Warehouse Receipt Header"; RelatedPO: Record "Purchase Header")
    var
        body: TextBuilder;
        html: Text;
        subj: Text;
        etaDate: Date;
        etaTxt: Text;
        Email: Codeunit Email;
        EmailMsg: Codeunit "Email Message";
        ToList: List of [Text];
        // for table build
        WRLine: Record "Warehouse Receipt Line";
        rowHtml: Text;
        anyRows: Boolean;
    begin
        // --- telemetry: start
        Session.LogMessage(
            'wr.shipped.start',
            StrSubstNo('SHIPPED: WR %1 for PO %2', WhseRcptHeader."No.", RelatedPO."No."),
            Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
            'WR', WhseRcptHeader."No.");

        // --- ETA resolution (Header.Expected Receipt Date, else max line expected)
        etaDate := GetPOETA(RelatedPO);
        if etaDate <> 0D then begin
            etaTxt := Format(etaDate);
            Session.LogMessage(
                'wr.shipped.eta.ok',
                StrSubstNo('ETA resolved = %1 for PO %2', etaTxt, RelatedPO."No."),
                Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
                'ETA', etaTxt);
        end else
            Session.LogMessage(
                'wr.shipped.eta.miss',
                StrSubstNo('No ETA found for PO %1 (header/lines empty)', RelatedPO."No."),
                Verbosity::Warning, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
                'PO', RelatedPO."No.");

        // --- subject
        subj := StrSubstNo('PO %1 Shipped', RelatedPO."No.");

        // --- body header
        body.AppendLine('<p>Dear Colleague,</p>');
        if etaTxt <> '' then
            body.AppendLine(StrSubstNo(
              '<p>PO <strong>#%1</strong> has <strong>Shipped!!</strong> The estimated arrival date is <strong>%2</strong>. The items and quantities are:</p>',
              Html(RelatedPO."No."), Html(etaTxt)))
        else
            body.AppendLine(StrSubstNo(
              '<p>PO <strong>#%1</strong> has <strong>Shipped!!</strong> The items and quantities are:</p>',
              Html(RelatedPO."No.")));

        // --- bordered table (only the PO’s lines on this WR)
        body.AppendLine(
          '<table style="border-collapse:collapse;margin:14px 0;width:100%;">' +
          '<tr style="background:#005fb8;color:#fff;">' +
          '<th style="padding:8px 10px;border:1px solid #d9d9d9;text-align:left;">Item No.</th>' +
          '<th style="padding:8px 10px;border:1px solid #d9d9d9;text-align:left;">Description</th>' +
          '<th style="padding:8px 10px;border:1px solid #d9d9d9;text-align:right;">Qty</th>' +
          '<th style="padding:8px 10px;border:1px solid #d9d9d9;text-align:left;">Location</th>' +
          '</tr>');

        WRLine.SetRange("No.", WhseRcptHeader."No.");
        WRLine.SetRange("Source Type", Database::"Purchase Line");
        WRLine.SetRange("Source No.", RelatedPO."No.");
        if WRLine.FindSet() then
            repeat
                anyRows := true;
                rowHtml :=
                  '<tr>' +
                  StrSubstNo('<td style="padding:8px 10px;border:1px solid #d9d9d9;">%1</td>', Html(WRLine."Item No.")) +
                  StrSubstNo('<td style="padding:8px 10px;border:1px solid #d9d9d9;">%1</td>', Html(WRLine.Description)) +
                  StrSubstNo('<td style="padding:8px 10px;border:1px solid #d9d9d9;text-align:right;">%1</td>', Format(WRLine.Quantity)) +
                  StrSubstNo('<td style="padding:8px 10px;border:1px solid #d9d9d9;">%1</td>', Html(WhseRcptHeader."Location Code")) +
                  '</tr>';
                body.AppendLine(rowHtml);
            until WRLine.Next() = 0;

        if not anyRows then begin
            // fallback line so the email doesn’t look empty if filtering failed
            body.AppendLine('<tr><td colspan="4" style="padding:10px;border:1px solid #d9d9d9;">No receipt lines found for this PO on the warehouse receipt.</td></tr>');
            Session.LogMessage(
                'wr.shipped.lines.miss',
                StrSubstNo('No WR lines matched WR %1 + PO %2', WhseRcptHeader."No.", RelatedPO."No."),
                Verbosity::Warning, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
                'WR', WhseRcptHeader."No.");
        end;

        body.AppendLine('</table>');
        body.AppendLine(BuildFooter());

        html := WrapHtml(subj, body.ToText());

        // recipients
        ToList := BuildRecipientListFromPO(RelatedPO);
        if ToList.Count() = 0 then begin
            Session.LogMessage(
                'wr.shipped.norecip',
                StrSubstNo('No recipients resolved for PO %1', RelatedPO."No."),
                Verbosity::Warning, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
                'PO', RelatedPO."No.");
            exit;
        end;

        EmailMsg.Create(ToList, subj, html, true);
        Email.Send(EmailMsg, Enum::"Email Scenario"::Default);

        Session.LogMessage(
            'wr.shipped.sent',
            StrSubstNo('SHIPPED email sent for WR %1 / PO %2', WhseRcptHeader."No.", RelatedPO."No."),
            Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
            'WR', WhseRcptHeader."No.");
    end;

    // Decide ETA for the PO: header Expected Receipt Date; if empty, max of line Expected Receipt Date.
    local procedure GetPOETA(var PO: Record "Purchase Header"): Date
    var
        POLine: Record "Purchase Line";
        best: Date;
    begin
        if PO."Expected Receipt Date" <> 0D then
            exit(PO."Expected Receipt Date");

        POLine.SetRange("Document Type", PO."Document Type");
        POLine.SetRange("Document No.", PO."No.");
        POLine.SetRange(Type, POLine.Type::Item);
        if POLine.FindSet() then
            repeat
                if POLine."Expected Receipt Date" > best then
                    best := POLine."Expected Receipt Date";
            until POLine.Next() = 0;

        exit(best); // 0D if still unknown
    end;

    // Helper: find the first non-blank ETA (header Expected -> header Promised -> earliest line date)
    local procedure GetPOEarliestETA(var PurchHeader: Record "Purchase Header"): Text
    var
        PurchLine: Record "Purchase Line";
        best: Date;
        d: Date;
    begin
        if PurchHeader."Expected Receipt Date" <> 0D then
            exit(Format(PurchHeader."Expected Receipt Date"));

        if PurchHeader."Promised Receipt Date" <> 0D then
            exit(Format(PurchHeader."Promised Receipt Date"));

        // Scan lines for earliest Expected/Promised
        best := DMY2DATE(31, 12, 9999);
        PurchLine.SetRange("Document Type", PurchHeader."Document Type");
        PurchLine.SetRange("Document No.", PurchHeader."No.");
        if PurchLine.FindSet() then
            repeat
                d := 0D;
                if PurchLine."Expected Receipt Date" <> 0D then
                    d := PurchLine."Expected Receipt Date"
                else
                    if PurchLine."Promised Receipt Date" <> 0D then
                        d := PurchLine."Promised Receipt Date";

                if (d <> 0D) and (d < best) then
                    best := d;
            until PurchLine.Next() = 0;

        if best <> DMY2DATE(31, 12, 9999) then
            exit(Format(best));

        exit('TBD');
    end;

    local procedure BuildHtmlTableFromPostedWRLines_All(PostedHdr: Record "Posted Whse. Receipt Header"; PONo: Code[20]): Text
    var
        PWRL: Record "Posted Whse. Receipt Line";
        t: TextBuilder;
    begin
        t.AppendLine('<table border="1" cellpadding="8" cellspacing="0" ' +
                     'style="border-collapse:collapse;width:100%;max-width:900px;margin:12px 0;border:1px solid #ccc;">');

        t.AppendLine('<tr style="background:#0a66c2;color:#fff;">' +
                     '<th align="left"  style="border:1px solid #ccc;">Item No.</th>' +
                     '<th align="left"  style="border:1px solid #ccc;">Description</th>' +
                     '<th align="right" style="border:1px solid #ccc;">Qty</th>' +
                     '<th align="left"  style="border:1px solid #ccc;">Location</th>' +
                     '</tr>');

        // Only lines for *this* posted receipt and *this* PO
        PWRL.SetRange("No.", PostedHdr."No.");
        PWRL.SetRange("Source Type", Database::"Purchase Line");
        PWRL.SetRange("Source No.", PONo);

        if PWRL.FindSet() then
            repeat
                if PWRL."Item No." <> '' then
                    t.AppendLine(StrSubstNo(
                        '<tr>' +
                        '<td style="border:1px solid #ccc;">%1</td>' +
                        '<td style="border:1px solid #ccc;">%2</td>' +
                        '<td style="border:1px solid #ccc;text-align:right;">%3</td>' +
                        '<td style="border:1px solid #ccc;">%4</td>' +
                        '</tr>',
                        Html(PWRL."Item No."),
                        Html(PWRL.Description),
                        Html(Format(PWRL.Quantity)), // use Quantity on posted line
                        Html(PWRL."Location Code")));
            until PWRL.Next() = 0;

        t.AppendLine('</table>');
        exit(t.ToText());
    end;

    procedure Notify_Arrived_OnWhseReceiptPosted(PostedHdr: Record "Posted Whse. Receipt Header"; RelatedPO: Record "Purchase Header")
    var
        subj: Text;
        body: TextBuilder;
        html: Text;
        Email: Codeunit Email;
        EmailMsg: Codeunit "Email Message";
        ToList: List of [Text];
        PostedLine: Record "Posted Whse. Receipt Line";
        LineCount: Integer;
    begin
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

        Session.LogMessage(
            'arrived.count',
            StrSubstNo('Arrived posted lines: %1 (PostedWR=%2, PO=%3)', LineCount, PostedHdr."No.", RelatedPO."No."),
            Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
            'Count', Format(LineCount));

        if LineCount = 0 then
            exit; // nothing to show, bail quietly

        body.AppendLine('<p>Dear Colleague,</p>');
        body.AppendLine(StrSubstNo(
            '<p>PO <strong>#%1</strong> has <strong>Arrived</strong> at our Bestway USA Chandler Warehouse!! ' +
            'Please allow 3–5 days for Container Unloading &amp; Putaways. The items and quantities on this shipment are:</p>',
            Html(RelatedPO."No.")));

        // Bordered table with all posted lines (right-aligned Qty)
        body.AppendLine(BuildHtmlTableFromPostedWRLines_All(PostedHdr, RelatedPO."No."));

        body.AppendLine('<p>If you have any questions, please contact your Supply Chain Team: ' +
                        '<a href="mailto:cyndy.peterson@bestwaycorp.us">cyndy.peterson@bestwaycorp.us</a> and/or ' +
                        '<a href="mailto:Eric.Eichstaedt@bestwaycorp.us">Eric.Eichstaedt@bestwaycorp.us</a>.</p>');

        html := WrapHtml(subj, body.ToText());

        // Recipients + send
        ToList := BuildRecipientListFromPO(RelatedPO);
        if ToList.Count() = 0 then
            exit;

        EmailMsg.Create(ToList, subj, html, true);
        Email.Send(EmailMsg, Enum::"Email Scenario"::Default);

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

    local procedure DispatchArrived(PostedNo: Code[20]; FromUnpostedNo: Code[20]; SourceTag: Text[10])
    var
        Helper: Codeunit "PO Email Helper";
        EmailState: Codeunit "PO Email State"; // ✅ add this so it's in scope
        PostedHdr: Record "Posted Whse. Receipt Header";
        PostedLine: Record "Posted Whse. Receipt Line";
        PurchHeader: Record "Purchase Header";
        PONumbers: List of [Code[20]];
        PONo: Code[20];
        RetryCount: Integer;
    begin
        if PostedNo = '' then
            exit;

        if not EmailState.ShouldSendArrived(PostedNo) then
            exit;

        if not PostedHdr.Get(PostedNo) then
            exit;

        // 🔄 Retry loop: wait for posted lines to be committed
        RetryCount := 0;
        repeat
            Clear(PostedLine);
            PostedLine.SetRange("No.", PostedNo);
            if PostedLine.FindFirst() then
                break;
            Sleep(1000); // wait 1 sec
            RetryCount += 1;
        until RetryCount > 10; // max 10s

        if not PostedLine.FindFirst() then begin
            Session.LogMessage(
                'wr.arrived.nolines',
                StrSubstNo('Posted WR %1 still has no lines after waiting.', PostedNo),
                Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
                'PostedWR', PostedNo);
            exit;
        end;

        // ✅ Collect distinct POs
        PostedLine.SetRange("No.", PostedNo);
        if PostedLine.FindSet() then
            repeat
                if PostedLine."Source Type" = Database::"Purchase Line" then begin
                    PONo := PostedLine."Source No.";
                    if (PONo <> '') and not PONumbers.Contains(PONo) then
                        PONumbers.Add(PONo);
                end;
            until PostedLine.Next() = 0;

        foreach PONo in PONumbers do
            if PurchHeader.Get(PurchHeader."Document Type"::Order, PONo) then
                Helper.Notify_Arrived_OnWhseReceiptPosted(PostedHdr, PurchHeader);

        EmailState.MarkArrivedSent(PostedNo);
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