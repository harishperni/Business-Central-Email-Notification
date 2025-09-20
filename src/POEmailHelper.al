codeunit 50142 "PO Email Helper"
{
    // =========================================================
    // PUBLIC ENTRY POINTS (called by subscribers)
    // =========================================================

    // PO Released -> "Created" (HTML body with table, no attachment)
    procedure Notify_POCreated_OnRelease(PurchHeader: Record "Purchase Header")
    var
        subj: Text;
        html: Text;
    begin
        if PurchHeader."Document Type" <> PurchHeader."Document Type"::Order then
            exit;

        subj := StrSubstNo('PO %1 Created', PurchHeader."No.");
        html := WrapHtml(subj, BuildHtmlBodyForPOCreated(PurchHeader));
        SendHtml(subj, html, BuildRecipientListFromPO(PurchHeader));
    end;

    // Warehouse Receipt created -> "Shipped" (attach THIS WR PDF only)
    procedure Notify_Shipped_OnWhseReceiptCreated(WhseRcptHeader: Record "Warehouse Receipt Header"; RelatedPO: Record "Purchase Header")
    var
        subj: Text;
        htmlBody: Text;
        eta: Text;
        attachName: Text[250];
        attachIn: InStream;
    begin
        if RelatedPO."Expected Receipt Date" <> 0D then
            eta := Format(RelatedPO."Expected Receipt Date");

        subj := StrSubstNo('PO %1 Shipped', RelatedPO."No.");

        if eta <> '' then
            htmlBody :=
              WrapHtml(subj,
                StrSubstNo(
                  '<p>Dear Colleague,</p>' +
                  '<p>PO <strong>#%1</strong> has <strong>Shipped!!</strong> The estimated arrival date is %2.</p>%3',
                  Html(RelatedPO."No."), Html(eta), BuildFooter()))
        else
            htmlBody :=
              WrapHtml(subj,
                StrSubstNo(
                  '<p>Dear Colleague,</p>' +
                  '<p>PO <strong>#%1</strong> has <strong>Shipped!!</strong>.</p>%2',
                  Html(RelatedPO."No."), BuildFooter()));

        // Render single-document PDF for THIS header only (Report 7316)
        AttachWrReceiptPdfStream(WhseRcptHeader, attachName, attachIn);

        SendHtmlWithAttachment(subj, htmlBody, BuildRecipientListFromPO(RelatedPO),
                               attachName, 'application/pdf', attachIn);
    end;

    // Posted WR created -> "Arrived" (attach THIS posted WR PDF only)
    procedure Notify_Arrived_OnWhseReceiptPosted(PostedWhseRcptHeader: Record "Posted Whse. Receipt Header"; RelatedPO: Record "Purchase Header")
    var
        subj: Text;
        htmlBody: Text;
        attachName: Text[250];
        attachIn: InStream;
    begin
        subj := StrSubstNo('PO %1 Arrived', RelatedPO."No.");

        htmlBody :=
          WrapHtml(subj,
            StrSubstNo(
              '<p>Dear Colleague,</p>' +
              '<p>PO <strong>#%1</strong> has <strong>Arrived</strong> at our Bestway USA Chandler Warehouse!! ' +
              'Please allow 3–5 days for Container Unloading &amp; Putaways.</p>%2',
              Html(RelatedPO."No."), BuildFooter()));

        // Render single-document PDF for THIS posted header only (Report 7308)
        AttachPostedWrReceiptPdfStream(PostedWhseRcptHeader, attachName, attachIn);

        SendHtmlWithAttachment(subj, htmlBody, BuildRecipientListFromPO(RelatedPO),
                               attachName, 'application/pdf', attachIn);
    end;


    // =========================================================
    // REPORT IDS (centralized)
    // =========================================================

    local procedure GetWhseReceiptReportId(): Integer
    begin
        // Standard "Warehouse Receipt" report
        exit(7316);
    end;

    local procedure GetPostedWhseReceiptReportId(): Integer
    begin
        // Standard "Warehouse Posted Receipt" report
        exit(7308);
    end;


    // =========================================================
    // RECIPIENTS (Contact first, then Vendor; no CCs)
    // =========================================================

    local procedure BuildRecipientListFromPO(PurchHeader: Record "Purchase Header"): List of [Text]
    var
        list: List of [Text];
        email: Text;
    begin
        email := ResolvePOPrimaryEmail(PurchHeader);
        if email <> '' then
            list.Add(email);
        exit(list);
    end;

    local procedure ResolvePOPrimaryEmail(PurchHeader: Record "Purchase Header"): Text
    var
        Contact: Record Contact;
        Vendor: Record Vendor;
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


    // =========================================================
    // EMAIL SEND (HTML + attachment using 5-arg AddAttachment)
    // =========================================================

    local procedure SendHtml(SubjectTxt: Text; HtmlBody: Text; ToList: List of [Text])
    var
        Email: Codeunit Email;
        Msg: Codeunit "Email Message";
        r: Text;
    begin
        if ToList.Count() = 0 then
            exit;

        Msg.Create('', SubjectTxt, HtmlBody, true);
        foreach r in ToList do
            Msg.AddRecipient(Enum::"Email Recipient Type"::"To", r);

        Email.Send(Msg, Enum::"Email Scenario"::Default);
    end;

    local procedure SendHtmlWithAttachment(SubjectTxt: Text; HtmlBody: Text; ToList: List of [Text]; AttachmentName: Text[250]; MimeType: Text[250]; AttachmentIn: InStream)
    var
        Email: Codeunit Email;
        Msg: Codeunit "Email Message";
        r: Text;
    begin
        if ToList.Count() = 0 then
            exit;

        Msg.Create('', SubjectTxt, HtmlBody, true);
        foreach r in ToList do
            Msg.AddRecipient(Enum::"Email Recipient Type"::"To", r);

        // BC v25 signature: AddAttachment(Name, MimeType, IsInline, ContentId, InStream)
        Msg.AddAttachment(AttachmentName, MimeType, false, '', AttachmentIn);

        Email.Send(Msg, Enum::"Email Scenario"::Default);
    end;


    // =========================================================
    // HTML BUILDERS
    // =========================================================

    local procedure BuildHtmlBodyForPOCreated(PurchHeader: Record "Purchase Header"): Text
    var
        tb: TextBuilder;
    begin
        tb.AppendLine('<p>Dear Colleague,</p>');
        tb.AppendLine(StrSubstNo('<p>PO <strong>#%1</strong> has been created per your request. You will receive another notification when your item(s) ship.</p>', Html(PurchHeader."No.")));
        tb.AppendLine('<p>The items and quantities are:</p>');
        tb.AppendLine(BuildHtmlTableFromPOLines(PurchHeader));
        tb.AppendLine(BuildFooter());
        exit(tb.ToText());
    end;

    local procedure BuildHtmlTableFromPOLines(PurchHeader: Record "Purchase Header"): Text
    var
        line: Record "Purchase Line";
        tb: TextBuilder;
    begin
        tb.AppendLine('<table style="border-collapse:collapse;border:1px solid #ddd;width:100%;font-family:Segoe UI,Arial,sans-serif;font-size:12px;">');
        tb.AppendLine('<thead><tr style="background:#0b5cab;color:#fff;">' +
                        '<th style="text-align:left;padding:6px;border:1px solid #ddd;">Item No.</th>' +
                        '<th style="text-align:left;padding:6px;border:1px solid #ddd;">Description</th>' +
                        '<th style="text-align:right;padding:6px;border:1px solid #ddd;">Qty</th>' +
                      '</tr></thead><tbody>');

        line.SetRange("Document Type", PurchHeader."Document Type");
        line.SetRange("Document No.", PurchHeader."No.");
        if line.FindSet() then
            repeat
                if (line.Type = line.Type::Item) and (line."No." <> '') then
                    tb.AppendLine(StrSubstNo(
                        '<tr><td style="padding:6px;border:1px solid #ddd;">%1</td>' +
                        '<td style="padding:6px;border:1px solid #ddd;">%2</td>' +
                        '<td style="padding:6px;border:1px solid #ddd;text-align:right;">%3</td></tr>',
                        Html(line."No."), Html(line.Description), Html(Format(line.Quantity))));
            until line.Next() = 0;

        tb.AppendLine('</tbody></table>');
        exit(tb.ToText());
    end;

    local procedure BuildFooter(): Text
    begin
        exit('<p>If you have any questions, please contact your Supply Chain Team: ' +
             '<a href="mailto:cyndy.peterson@bestwaycorp.us">cyndy.peterson@bestwaycorp.us</a> and/or ' +
             '<a href="mailto:Eric.Eichstaedt@bestwaycorp.us">Eric.Eichstaedt@bestwaycorp.us</a>.</p>');
    end;

    local procedure Html(Value: Text): Text
    begin
        // NOTE: ConvertStr can't output multi-character entities.
        // To avoid runtime errors during posting/release, keep this a no-op.
        exit(Value);
    end;

    local procedure WrapHtml(Title: Text; BodyInner: Text): Text
    var
        tb: TextBuilder;
    begin
        tb.AppendLine('<!DOCTYPE html><html><head><meta charset="utf-8"><title>' + Html(Title) + '</title></head>');
        tb.AppendLine('<body style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;line-height:1.35;">');
        tb.AppendLine(BodyInner);
        tb.AppendLine('</body></html>');
        exit(tb.ToText());
    end;


    // =========================================================
    // PDF GENERATION (single-doc; filtered; RecordRef overload)
    // =========================================================

    local procedure AttachWrReceiptPdfStream(WhseHdr: Record "Warehouse Receipt Header"; var FileName: Text[250]; var InS: InStream)
    var
        TmpBlob: Codeunit "Temp Blob";
        OutS: OutStream;
        H: Record "Warehouse Receipt Header";
        RRef: RecordRef;
    begin
        if WhseHdr."No." = '' then
            Error('Warehouse Receipt Header has no No.');

        // Ensure single-document scope
        H.Reset();
        H.SetRange("No.", WhseHdr."No.");

        FileName := StrSubstNo('WarehouseReceipt_%1.pdf', H."No.");

        TmpBlob.CreateOutStream(OutS);
        RRef.GetTable(H);
        Report.SaveAs(GetWhseReceiptReportId(), '', ReportFormat::Pdf, OutS, RRef);
        TmpBlob.CreateInStream(InS);
    end;

    local procedure AttachPostedWrReceiptPdfStream(PostedHdr: Record "Posted Whse. Receipt Header"; var FileName: Text[250]; var InS: InStream)
    var
        TmpBlob: Codeunit "Temp Blob";
        OutS: OutStream;
        H: Record "Posted Whse. Receipt Header";
        RRef: RecordRef;
    begin
        if PostedHdr."No." = '' then
            Error('Posted Whse. Receipt Header has no No.');

        // Ensure single-document scope
        H.Reset();
        H.SetRange("No.", PostedHdr."No.");

        FileName := StrSubstNo('PostedWarehouseReceipt_%1.pdf', H."No.");

        TmpBlob.CreateOutStream(OutS);
        RRef.GetTable(H);
        Report.SaveAs(GetPostedWhseReceiptReportId(), '', ReportFormat::Pdf, OutS, RRef);
        TmpBlob.CreateInStream(InS);
    end;
}