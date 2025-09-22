codeunit 50142 "PO Email Helper"
{
    procedure Notify_POCreated_OnRelease(PurchHeader: Record "Purchase Header")
    var
        htmlBody: Text;
        subject: Text;
        toList: List of [Text];
    begin
        if not EnsureRecipientsFromPO(PurchHeader, toList) then
            exit;

        subject := StrSubstNo('PO %1 Created', PurchHeader."No.");
        htmlBody := WrapHtml(
            subject,
            BuildGreeting() +
            StrSubstNo('<p>PO <strong>#%1</strong> has been created per your request.</p>', Html(PurchHeader."No.")) +
            '<p>You will receive another notification when your item(s) ship. The items and quantities are:</p>' +
            BuildHtmlTableFromPOLines(PurchHeader) +
            BuildSignature());

        SendEmail(subject, htmlBody, toList);
    end;

    procedure Notify_Shipped_OnWhseReceiptCreated(WhseRcptHeader: Record "Warehouse Receipt Header"; RelatedPO: Record "Purchase Header")
    var
        htmlBody: Text;
        subject: Text;
        toList: List of [Text];
        etaTxt: Text;
    begin
        if not EnsureRecipientsFromPO(RelatedPO, toList) then
            exit;

        if RelatedPO."Expected Receipt Date" <> 0D then
            etaTxt := Format(RelatedPO."Expected Receipt Date");

        subject := StrSubstNo('PO %1 Shipped', RelatedPO."No.");

        htmlBody := BuildGreeting();
        if etaTxt <> '' then
            htmlBody := htmlBody + StrSubstNo(
                '<p>PO <strong>#%1</strong> has <strong>Shipped!!</strong> with estimated arrival date %2.</p>',
                Html(RelatedPO."No."), Html(etaTxt))
        else
            htmlBody := htmlBody + StrSubstNo(
                '<p>PO <strong>#%1</strong> has <strong>Shipped!!</strong>.</p>',
                Html(RelatedPO."No."));

        htmlBody := htmlBody + '<p>The items and quantities on this shipment are:</p>';
        htmlBody := htmlBody + BuildHtmlTableFromWRLines(WhseRcptHeader);
        htmlBody := htmlBody + BuildSignature();
        htmlBody := WrapHtml(subject, htmlBody);

        SendEmail(subject, htmlBody, toList);
    end;

    procedure Notify_Arrived_OnWhseReceiptPosted(PostedHdr: Record "Posted Whse. Receipt Header"; RelatedPO: Record "Purchase Header")
    var
        htmlBody: Text;
        subject: Text;
        toList: List of [Text];
    begin
        if not EnsureRecipientsFromPO(RelatedPO, toList) then
            exit;

        subject := StrSubstNo('PO %1 Arrived', RelatedPO."No.");
        htmlBody := WrapHtml(
            subject,
            BuildGreeting() +
            StrSubstNo('<p>PO <strong>#%1</strong> has <strong>Arrived</strong> at our warehouse. The items and quantities on this receipt are:</p>', Html(RelatedPO."No.")) +
            BuildHtmlTableFromPostedWRLines(PostedHdr) +
            BuildSignature());

        SendEmail(subject, htmlBody, toList);
    end;

    // ---------------------- recipients ----------------------
    local procedure EnsureRecipientsFromPO(PurchHeader: Record "Purchase Header"; var ToList: List of [Text]): Boolean
    var
        primary: Text;
    begin
        Clear(ToList);

        primary := ResolvePOPrimaryEmail(PurchHeader);
        if primary <> '' then
            ToList.Add(primary);

        exit(ToList.Count() > 0);
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

    // ---------------------- HTML ----------------------
    local procedure BuildGreeting(): Text
    begin
        exit('<p>Dear Colleague,</p>');
    end;

    local procedure BuildSignature(): Text
    begin
        exit('<p>If you have any questions, please contact your Supply Chain Team: ' +
             '<a href="mailto:cyndy.peterson@bestwaycorp.us">cyndy.peterson@bestwaycorp.us</a> and/or ' +
             '<a href="mailto:Eric.Eichstaedt@bestwaycorp.us">Eric.Eichstaedt@bestwaycorp.us</a>.</p>');
    end;

    local procedure BuildHtmlTableFromPOLines(PurchHeader: Record "Purchase Header"): Text
    var
        PurchLine: Record "Purchase Line";
        b: TextBuilder;
    begin
        b.AppendLine('<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">');
        b.AppendLine('<thead><tr><th align="left">Item</th><th align="left">Description</th><th align="right">Quantity</th></tr></thead>');
        b.AppendLine('<tbody>');

        PurchLine.SetRange("Document Type", PurchHeader."Document Type");
        PurchLine.SetRange("Document No.", PurchHeader."No.");
        if PurchLine.FindSet() then
            repeat
                if (PurchLine.Type = PurchLine.Type::Item) and (PurchLine."No." <> '') then
                    b.AppendLine(StrSubstNo(
                        '<tr><td>%1</td><td>%2</td><td align="right">%3</td></tr>',
                        Html(PurchLine."No."),
                        Html(PurchLine.Description),
                        Html(Format(PurchLine.Quantity))));
            until PurchLine.Next() = 0;

        b.AppendLine('</tbody></table>');
        exit(b.ToText());
    end;

    local procedure BuildHtmlTableFromWRLines(WhseRcptHeader: Record "Warehouse Receipt Header"): Text
    var
        WhseRcptLine: Record "Warehouse Receipt Line";
        b: TextBuilder;
    begin
        b.AppendLine('<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">');
        b.AppendLine('<thead><tr><th align="left">Item</th><th align="left">Description</th><th align="right">Qty. to Receive</th></tr></thead>');
        b.AppendLine('<tbody>');

        WhseRcptLine.SetRange("No.", WhseRcptHeader."No.");
        if WhseRcptLine.FindSet() then
            repeat
                if WhseRcptLine."Item No." <> '' then
                    b.AppendLine(StrSubstNo(
                        '<tr><td>%1</td><td>%2</td><td align="right">%3</td></tr>',
                        Html(WhseRcptLine."Item No."),
                        Html(WhseRcptLine.Description),
                        Html(Format(WhseRcptLine."Qty. to Receive"))));
            until WhseRcptLine.Next() = 0;

        b.AppendLine('</tbody></table>');
        exit(b.ToText());
    end;

    local procedure BuildHtmlTableFromPostedWRLines(PostedHdr: Record "Posted Whse. Receipt Header"): Text
    var
        PostedLine: Record "Posted Whse. Receipt Line";
        b: TextBuilder;
    begin
        b.AppendLine('<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">');
        b.AppendLine('<thead><tr><th align="left">Item</th><th align="left">Description</th><th align="right">Qty.</th></tr></thead>');
        b.AppendLine('<tbody>');

        PostedLine.SetRange("No.", PostedHdr."No.");
        if PostedLine.FindSet() then
            repeat
                if PostedLine."Item No." <> '' then
                    b.AppendLine(StrSubstNo(
                        '<tr><td>%1</td><td>%2</td><td align="right">%3</td></tr>',
                        Html(PostedLine."Item No."),
                        Html(PostedLine.Description),
                        Html(Format(PostedLine.Quantity))));
            until PostedLine.Next() = 0;

        b.AppendLine('</tbody></table>');
        exit(b.ToText());
    end;

    local procedure WrapHtml(title: Text; inner: Text): Text
    begin
        exit(
          '<!DOCTYPE html><html><head><meta charset="utf-8">' +
          '<title>' + Html(title) + '</title></head><body style="font-family:Segoe UI,Arial,sans-serif;font-size:12pt;line-height:1.35">' +
          inner +
          '</body></html>');
    end;

    local procedure Html(t: Text): Text
    begin
        t := ReplaceText(t, '&', '&amp;');
        t := ReplaceText(t, '<', '&lt;');
        t := ReplaceText(t, '>', '&gt;');
        t := ReplaceText(t, '"', '&quot;');
        exit(t);
    end;

    local procedure ReplaceText(S: Text; Find: Text; NewText: Text): Text
    var
        P: Integer;
        R: Text;
        Lf: Integer;
    begin
        R := S;
        Lf := StrLen(Find);
        if Lf = 0 then
            exit(R);

        P := StrPos(R, Find);
        while P > 0 do begin
            R := CopyStr(R, 1, P - 1) + NewText + CopyStr(R, P + Lf);
            P := StrPos(R, Find);
        end;
        exit(R);
    end;

    // ---------------------- SEND ----------------------
    local procedure SendEmail(SubjectTxt: Text; HtmlBody: Text; ToList: List of [Text])
    var
        Email: Codeunit Email;
        EmailMessage: Codeunit "Email Message";
        addr: Text;
    begin
        if ToList.Count() = 0 then
            exit;

        EmailMessage.Create('', SubjectTxt, HtmlBody, true);
        foreach addr in ToList do
            EmailMessage.AddRecipient(Enum::"Email Recipient Type"::"To", addr);

        Email.Send(EmailMessage, Enum::"Email Scenario"::Default);
    end;
}