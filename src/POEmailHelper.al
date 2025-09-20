codeunit 50142 "PO Email Helper"
{
    // =========================================================
    // PUBLIC ENTRY POINTS
    // =========================================================

    // A) PO Created: HTML table (no attachment)
    procedure Notify_POCreated_OnRelease(PurchHeader: Record "Purchase Header")
    var
        subj: Text;
        html: Text;
    begin
        if PurchHeader."Document Type" <> PurchHeader."Document Type"::Order then
            exit;

        subj := StrSubstNo('PO %1 Created', PurchHeader."No.");
        html := WrapHtml(subj, BuildHtmlForPOCreated(PurchHeader));
        SendHtml(subj, html, BuildRecipientListFromPO(PurchHeader));
    end;

    // B) Shipped: HTML table with ONE line from this WR (no attachment)
    procedure Notify_Shipped_OnWhseReceiptCreated(WhseRcptHeader: Record "Warehouse Receipt Header"; RelatedPO: Record "Purchase Header")
    var
        subj: Text;
        body: Text;
        html: Text;
        eta: Text;
    begin
        if RelatedPO."Expected Receipt Date" <> 0D then
            eta := Format(RelatedPO."Expected Receipt Date");

        body += '<p>Dear Colleague,</p>';
        if eta <> '' then
            body += StrSubstNo('<p>PO <strong>#%1</strong> has <strong>Shipped!!</strong> The estimated arrival date is %2.</p>',
                               Html(RelatedPO."No."), Html(eta))
        else
            body += StrSubstNo('<p>PO <strong>#%1</strong> has <strong>Shipped!!</strong>.</p>',
                               Html(RelatedPO."No."));
        body += '<p>Item shipped on this receipt:</p>';
        body += BuildHtmlForSingleWrLine(WhseRcptHeader);
        body += BuildFooter();

        subj := StrSubstNo('PO %1 Shipped', RelatedPO."No.");
        html := WrapHtml(subj, body);

        SendHtml(subj, html, BuildRecipientListFromPO(RelatedPO));
    end;

    // C) Arrived: attach the Posted Warehouse Receipt PDF (single document)
    procedure Notify_Arrived_OnWhseReceiptPosted(PostedWhseRcptHeader: Record "Posted Whse. Receipt Header"; RelatedPO: Record "Purchase Header")
    var
        subj: Text;
        body: Text;
        html: Text;
        attName: Text[250];
        attIn: InStream;
    begin
        subj := StrSubstNo('PO %1 Arrived', RelatedPO."No.");

        body += '<p>Dear Colleague,</p>';
        body += StrSubstNo(
            '<p>PO <strong>#%1</strong> has <strong>Arrived</strong> at our Bestway USA Chandler Warehouse!! ' +
            'Please allow 3–5 days for Container Unloading &amp; Putaways.</p>',
            Html(RelatedPO."No."));
        body += BuildFooter();

        html := WrapHtml(subj, body);

        // Create single-doc PDF for THIS posted WR (standard report 7308)
        AttachPostedWrReceiptPdf(PostedWhseRcptHeader, attName, attIn);

        SendHtmlWithAttachment(subj, html, BuildRecipientListFromPO(RelatedPO),
                               attName, 'application/pdf', attIn);
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

    local procedure ResolvePOPrimaryEmail(PH: Record "Purchase Header"): Text
    var
        Contact: Record Contact;
        Vendor: Record Vendor;
    begin
        if PH."Buy-from Contact No." <> '' then
            if Contact.Get(PH."Buy-from Contact No.") then
                if Contact."E-Mail" <> '' then
                    exit(Contact."E-Mail");

        if Vendor.Get(PH."Buy-from Vendor No.") then
            if Vendor."E-Mail" <> '' then
                exit(Vendor."E-Mail");

        exit('');
    end;


    // =========================================================
    // SENDING (no risky encoding; 5-arg AddAttachment)
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

        Msg.AddAttachment(AttachmentName, MimeType, false, '', AttachmentIn);
        Email.Send(Msg, Enum::"Email Scenario"::Default);
    end;


    // =========================================================
    // HTML BUILDERS
    // =========================================================

    local procedure BuildHtmlForPOCreated(PH: Record "Purchase Header"): Text
    var
        b: TextBuilder;
        line: Record "Purchase Line";
    begin
        b.AppendLine(StrSubstNo('<p>PO <strong>#%1</strong> has been created per your request. You will receive another notification when your item(s) ship.</p>', Html(PH."No.")));
        b.AppendLine('<p>The items and quantities are:</p>');

        b.AppendLine('<table style="border-collapse:collapse;border:1px solid #ddd;width:100%;font-family:Segoe UI,Arial,sans-serif;font-size:12px;">');
        b.AppendLine('<thead><tr style="background:#0b5cab;color:#fff;">' +
                        '<th style="text-align:left;padding:6px;border:1px solid #ddd;">Item No.</th>' +
                        '<th style="text-align:left;padding:6px;border:1px solid #ddd;">Description</th>' +
                        '<th style="text-align:right;padding:6px;border:1px solid #ddd;">Qty</th>' +
                      '</tr></thead><tbody>');

        line.SetRange("Document Type", PH."Document Type");
        line.SetRange("Document No.", PH."No.");
        if line.FindSet() then
            repeat
                if (line.Type = line.Type::Item) and (line."No." <> '') then
                    b.AppendLine(StrSubstNo(
                        '<tr><td style="padding:6px;border:1px solid #ddd;">%1</td>' +
                        '<td style="padding:6px;border:1px solid #ddd;">%2</td>' +
                        '<td style="padding:6px;border:1px solid #ddd;text-align:right;">%3</td></tr>',
                        Html(line."No."), Html(line.Description), Html(Format(line.Quantity))));
            until line.Next() = 0;

        b.AppendLine('</tbody></table>');
        b.AppendLine(BuildFooter());
        exit(b.ToText());
    end;

    // ONE WR line (first line on this receipt)
    local procedure BuildHtmlForSingleWrLine(WRHdr: Record "Warehouse Receipt Header"): Text
    var
        L: Record "Warehouse Receipt Line";
        b: TextBuilder;
        found: Boolean;
    begin
        b.AppendLine('<table style="border-collapse:collapse;border:1px solid #ddd;width:100%;font-family:Segoe UI,Arial,sans-serif;font-size:12px;">');
        b.AppendLine('<thead><tr style="background:#0b5cab;color:#fff;">' +
                        '<th style="text-align:left;padding:6px;border:1px solid #ddd;">Item No.</th>' +
                        '<th style="text-align:left;padding:6px;border:1px solid #ddd;">Description</th>' +
                        '<th style="text-align:right;padding:6px;border:1px solid #ddd;">Qty to Receive</th>' +
                      '</tr></thead><tbody>');

        L.SetRange("No.", WRHdr."No.");
        L.SetRange("Source Document", L."Source Document"::"Purchase Order");
        if L.FindFirst() then begin
            if L."Item No." <> '' then begin
                found := true;
                b.AppendLine(StrSubstNo(
                    '<tr><td style="padding:6px;border:1px solid #ddd;">%1</td>' +
                    '<td style="padding:6px;border:1px solid #ddd;">%2</td>' +
                    '<td style="padding:6px;border:1px solid #ddd;text-align:right;">%3</td></tr>',
                    Html(L."Item No."), Html(L.Description), Html(Format(L."Qty. to Receive"))));
            end;
        end;

        if not found then
            b.AppendLine('<tr><td colspan="3" style="padding:6px;border:1px solid #ddd;">No item lines found.</td></tr>');

        b.AppendLine('</tbody></table>');
        exit(b.ToText());
    end;

    local procedure BuildFooter(): Text
    begin
        exit('<p>If you have any questions, please contact your Supply Chain Team: ' +
             '<a href="mailto:cyndy.peterson@bestwaycorp.us">cyndy.peterson@bestwaycorp.us</a> and/or ' +
             '<a href="mailto:Eric.Eichstaedt@bestwaycorp.us">Eric.Eichstaedt@bestwaycorp.us</a>.</p>');
    end;

    local procedure Html(Value: Text): Text
    begin
        // IMPORTANT: keep this a no-op to avoid ConvertStr multi-char mapping runtime errors.
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
    // PDF ATTACHMENT (Posted WR only; single doc)
    // =========================================================

    local procedure AttachPostedWrReceiptPdf(PostedHdr: Record "Posted Whse. Receipt Header"; var FileName: Text[250]; var InS: InStream)
    var
        TmpBlob: Codeunit "Temp Blob";
        OutS: OutStream;
        H: Record "Posted Whse. Receipt Header";
        RRef: RecordRef;
        ReportId_PostedWhseReceipt: Integer;
    begin
        if PostedHdr."No." = '' then
            exit;

        // Standard Posted Warehouse Receipt report
        ReportId_PostedWhseReceipt := 7308;

        // Filter to THIS posted document only
        H.Reset();
        H.SetRange("No.", PostedHdr."No.");

        FileName := StrSubstNo('PostedWarehouseReceipt_%1.pdf', H."No.");

        TmpBlob.CreateOutStream(OutS);
        RRef.GetTable(H);
        // Use overload: ReportId, '', ReportFormat, OutStream, RecordRef
        Report.SaveAs(ReportId_PostedWhseReceipt, '', ReportFormat::Pdf, OutS, RRef);
        TmpBlob.CreateInStream(InS);
    end;
}