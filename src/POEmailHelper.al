codeunit 50142 "PO Email Helper"
{
    // ==========================================
    // Public entry points
    // ==========================================

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
          PurchHeader."No."));
        body.AppendLine(BuildHtmlTableFromPOLines(PurchHeader));
        body.AppendLine(BuildFooter());

        html := WrapHtml(subj, body.ToText());
        SendEmailHtml(subj, html, BuildRecipientListFromPO(PurchHeader));
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
                '<p>PO <strong>#%1</strong> has <strong>Shipped!!</strong> The estimated arrival date is %2. The items and quantities are:</p>',
                RelatedPO."No.", eta))
        else
            body.AppendLine(StrSubstNo(
                '<p>PO <strong>#%1</strong> has <strong>Shipped!!</strong>. The items and quantities are:</p>',
                RelatedPO."No."));

        // WR lines for THIS PO on THIS WR
        body.AppendLine(BuildHtmlTableFromWhseRcptLines(WhseRcptHeader, RelatedPO));
        body.AppendLine(BuildFooter());

        html := WrapHtml(subj, body.ToText());
        SendEmailHtml(subj, html, BuildRecipientListFromPO(RelatedPO));
    end;

    procedure Notify_Arrived_OnWhseReceiptPosted(PostedHdr: Record "Posted Whse. Receipt Header"; RelatedPO: Record "Purchase Header")
    var
        subj: Text;
        body: TextBuilder;
        html: Text;
    begin
        subj := StrSubstNo('PO %1 Arrived', RelatedPO."No.");

        body.AppendLine('<p>Dear Colleague,</p>');
        body.AppendLine(StrSubstNo(
            '<p>PO <strong>#%1</strong> has <strong>Arrived</strong> at our Bestway USA Chandler Warehouse!! Please allow 3–5 days for Container Unloading &amp; Putaways. The items and quantities on this shipment are:</p>',
            RelatedPO."No."));

        // ALL posted lines for THIS posted receipt & THIS PO
        body.AppendLine(BuildHtmlTableFromPostedWhseRcptLines(PostedHdr, RelatedPO));
        body.AppendLine(BuildFooter());

        html := WrapHtml(subj, body.ToText());
        SendEmailHtml(subj, html, BuildRecipientListFromPO(RelatedPO));
    end;

    // ==========================================
    // Recipients (no CCs)
    // ==========================================

    local procedure BuildRecipientListFromPO(PurchHeader: Record "Purchase Header"): List of [Text]
    var
        list: List of [Text];
        primary: Text;
    begin
        primary := ResolvePOPrimaryEmail(PurchHeader);
        if primary <> '' then
            list.Add(primary);
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

    // ==========================================
    // HTML wrappers
    // ==========================================

    local procedure WrapHtml(Title: Text; BodyInner: Text): Text
    var
        tb: TextBuilder;
    begin
        tb.AppendLine('<!DOCTYPE html><html><head><meta charset="utf-8" />');
        tb.AppendLine('<style>');
        tb.AppendLine('table{border-collapse:collapse;width:100%;font-family:Segoe UI,Arial,sans-serif;font-size:12px}');
        tb.AppendLine('th,td{border:1px solid #d0d7de;padding:6px 8px;text-align:left}');
        tb.AppendLine('th{background:#0f6cbd;color:#fff}');
        tb.AppendLine('</style></head><body>');
        tb.AppendLine(StrSubstNo('<h3>%1</h3>', Title));
        tb.AppendLine(BodyInner);
        tb.AppendLine('</body></html>');
        exit(tb.ToText());
    end;

    local procedure BuildFooter(): Text
    begin
        exit(
          '<p>If you have any questions, please contact your Supply Chain Team: ' +
          '<a href="mailto:cyndy.peterson@bestwaycorp.us">cyndy.peterson@bestwaycorp.us</a> ' +
          'and/or <a href="mailto:Eric.Eichstaedt@bestwaycorp.us">Eric.Eichstaedt@bestwaycorp.us</a>.</p>');
    end;

    // ==========================================
    // Table builders
    // ==========================================

    local procedure BuildHtmlTableFromPOLines(PurchHeader: Record "Purchase Header"): Text
    var
        line: Record "Purchase Line";
        tb: TextBuilder;
    begin
        tb.AppendLine('<table><thead><tr>');
        tb.AppendLine('<th>Item No.</th><th>Description</th><th>Qty</th><th>Location</th>');
        tb.AppendLine('</tr></thead><tbody>');

        line.SetRange("Document Type", PurchHeader."Document Type");
        line.SetRange("Document No.", PurchHeader."No.");
        if line.FindSet() then
            repeat
                if (line.Type = line.Type::Item) and (line."No." <> '') then begin
                    tb.AppendLine('<tr>');
                    tb.AppendLine(StrSubstNo('<td>%1</td>', line."No."));
                    tb.AppendLine(StrSubstNo('<td>%1</td>', line.Description));
                    tb.AppendLine(StrSubstNo('<td style="text-align:right">%1</td>', Format(line.Quantity)));
                    tb.AppendLine(StrSubstNo('<td>%1</td>', line."Location Code"));
                    tb.AppendLine('</tr>');
                end;
            until line.Next() = 0;

        tb.AppendLine('</tbody></table>');
        exit(tb.ToText());
    end;

    local procedure BuildHtmlTableFromWhseRcptLines(WhseRcptHeader: Record "Warehouse Receipt Header"; OnlyPO: Record "Purchase Header"): Text
    var
        L: Record "Warehouse Receipt Line"; // 7317
        b: TextBuilder;
        hadAny: Boolean;
    begin
        b.AppendLine('<table><thead><tr>');
        b.AppendLine('<th>Item No.</th><th>Description</th><th>Qty</th><th>Location</th>');
        b.AppendLine('</tr></thead><tbody>');

        // Filter lines to THIS WR and THIS PO
        L.SetRange("No.", WhseRcptHeader."No.");
        L.SetRange("Source Document", L."Source Document"::"Purchase Order");
        L.SetRange("Source No.", OnlyPO."No.");

        if L.FindSet() then
            repeat
                hadAny := true;
                if (L."Item No." <> '') then begin
                    b.AppendLine('<tr>');
                    b.AppendLine(StrSubstNo('<td>%1</td>', L."Item No."));
                    b.AppendLine(StrSubstNo('<td>%1</td>', L.Description));
                    b.AppendLine(StrSubstNo('<td style="text-align:right">%1</td>', Format(L.Quantity)));
                    b.AppendLine(StrSubstNo('<td>%1</td>', L."Location Code"));
                    b.AppendLine('</tr>');
                end;
            until L.Next() = 0;

        if not hadAny then
            b.AppendLine('<tr><td colspan="4">No lines for this PO on this warehouse receipt.</td></tr>');

        b.AppendLine('</tbody></table>');
        exit(b.ToText());
    end;

    // ALL posted lines for THIS posted receipt & THIS PO (fix for "only 1 line" issue)
    local procedure BuildHtmlTableFromPostedWhseRcptLines(PostedHdr: Record "Posted Whse. Receipt Header"; OnlyPO: Record "Purchase Header"): Text
    var
        L: Record "Posted Whse. Receipt Line"; // 7319
        b: TextBuilder;
        hadAny: Boolean;
    begin
        b.AppendLine('<table><thead><tr>');
        b.AppendLine('<th>Item No.</th><th>Description</th><th>Qty</th><th>Location</th>');
        b.AppendLine('</tr></thead><tbody>');

        L.Reset();
        L.SetRange("No.", PostedHdr."No.");
        L.SetRange("Source Document", L."Source Document"::"Purchase Order");
        L.SetRange("Source No.", OnlyPO."No.");

        if L.FindSet() then
            repeat
                hadAny := true;
                b.AppendLine('<tr>');
                b.AppendLine(StrSubstNo('<td>%1</td>', L."Item No."));
                b.AppendLine(StrSubstNo('<td>%1</td>', L.Description));
                b.AppendLine(StrSubstNo('<td style="text-align:right">%1</td>', Format(L.Quantity)));
                b.AppendLine(StrSubstNo('<td>%1</td>', L."Location Code"));
                b.AppendLine('</tr>');
            until L.Next() = 0;

        if not hadAny then
            b.AppendLine('<tr><td colspan="4">No posted lines found for this PO on this receipt.</td></tr>');

        b.AppendLine('</tbody></table>');
        exit(b.ToText());
    end;

    // ==========================================
    // Email send (HTML)
    // ==========================================

    local procedure SendEmailHtml(SubjectTxt: Text; HtmlBody: Text; ToList: List of [Text])
    var
        Email: Codeunit Email;
        EmailMessage: Codeunit "Email Message";
    begin
        if ToList.Count() = 0 then
            exit;

        EmailMessage.Create(ToList, SubjectTxt, HtmlBody, true);
        Email.Send(EmailMessage, Enum::"Email Scenario"::Default);
    end;
}