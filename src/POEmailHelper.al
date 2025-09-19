codeunit 50142 "PO Email Helper"
{
    // ========================
    // PUBLIC NOTIFICATIONS
    // ========================

    procedure Notify_POCreated_OnRelease(PurchHeader: Record "Purchase Header")
    var
        subj: Text;
        html: Text;
        body: TextBuilder;
        tableHtml: Text;
    begin
        tableHtml := BuildHtmlTableFromPOLines(PurchHeader);

        body.AppendLine('<p>Dear Colleague,</p>');
        body.AppendLine(
          StrSubstNo('<p>PO <strong>#%1</strong> has been created per your request. You will receive another notification when your item(s) ship. The items and quantities are:</p>',
                     HtmlEncode(PurchHeader."No.")));
        body.AppendLine(tableHtml);
        body.AppendLine(SignatureHtml());

        subj := StrSubstNo('PO %1 Created', PurchHeader."No.");
        html := WrapHtml(subj, body.ToText());
        SendEmail(subj, html, BuildRecipientListFromPO(PurchHeader));
    end;

    procedure Notify_Shipped_OnWhseReceiptCreated(WhseRcptHeader: Record "Warehouse Receipt Header"; RelatedPO: Record "Purchase Header")
    var
        subj: Text;
        html: Text;
        body: TextBuilder;
        tableHtml: Text;
        eta: Text;
    begin
        eta := GetETAText(RelatedPO);
        tableHtml := BuildHtmlTableFromWhseRcptLines(WhseRcptHeader, RelatedPO);

        body.AppendLine('<p>Dear Colleague,</p>');
        if eta <> '' then
            body.AppendLine(StrSubstNo(
                '<p>PO <strong>#%1</strong> has <strong>Shipped!!</strong> The estimated arrival date is <strong>%2</strong>. The items and quantities are:</p>',
                HtmlEncode(RelatedPO."No."), HtmlEncode(eta)))
        else
            body.AppendLine(StrSubstNo(
                '<p>PO <strong>#%1</strong> has <strong>Shipped!!</strong>. The items and quantities are:</p>',
                HtmlEncode(RelatedPO."No.")));

        body.AppendLine(tableHtml);
        body.AppendLine(SignatureHtml());

        subj := StrSubstNo('PO %1 Shipped', RelatedPO."No.");
        html := WrapHtml(subj, body.ToText());
        SendEmail(subj, html, BuildRecipientListFromPO(RelatedPO));
    end;

    procedure Notify_Arrived_FromPostedReceipt(PostedWhseHdr: Record "Posted Whse. Receipt Header"; RelatedPO: Record "Purchase Header")
    var
        subj: Text;
        html: Text;
        body: TextBuilder;
        tableHtml: Text;
    begin
        tableHtml := BuildHtmlTableFromPostedWhseRcptLines(PostedWhseHdr, RelatedPO);

        body.AppendLine('<p>Dear Colleague,</p>');
        body.AppendLine(StrSubstNo(
          '<p>PO <strong>#%1</strong> has <strong>Arrived</strong> at our Bestway USA Chandler Warehouse!! Please allow 3–5 days for Container Unloading &amp; Putaways. The items and quantities on this shipment are:</p>',
          HtmlEncode(RelatedPO."No.")));
        body.AppendLine(tableHtml);
        body.AppendLine(SignatureHtml());

        subj := StrSubstNo('PO %1 Arrived', RelatedPO."No.");
        html := WrapHtml(subj, body.ToText());
        SendEmail(subj, html, BuildRecipientListFromPO(RelatedPO));
    end;

    // ========================
    // TABLE BUILDERS
    // ========================

    // PO Created → show ALL PO lines
    local procedure BuildHtmlTableFromPOLines(PH: Record "Purchase Header"): Text
    var
        L: Record "Purchase Line";
        tb: TextBuilder;
    begin
        tb.AppendLine(TableHeadHtml());
        L.SetRange("Document Type", PH."Document Type");
        L.SetRange("Document No.", PH."No.");
        if L.FindSet() then
            repeat
                if (L.Type = L.Type::Item) and (L."No." <> '') then
                    tb.AppendLine(TableRowHtml(L."No.", L.Description, L.Quantity, PH."Location Code"));
            until L.Next() = 0;
        tb.AppendLine('</tbody></table>');
        exit(tb.ToText());
    end;

    // Shipped → ONLY lines on this Warehouse Receipt for THIS PO
    local procedure BuildHtmlTableFromWhseRcptLines(H: Record "Warehouse Receipt Header"; RelatedPO: Record "Purchase Header"): Text
    var
        L: Record "Warehouse Receipt Line";
        tb: TextBuilder;
        loc: Code[10];
    begin
        tb.AppendLine(TableHeadHtml());
        L.Reset();
        L.SetRange("No.", H."No.");
        L.SetRange("Source Document", L."Source Document"::"Purchase Order");
        L.SetRange("Source No.", RelatedPO."No.");

        if L.FindSet() then
            repeat
                if L."Item No." <> '' then begin
                    loc := L."Location Code";
                    if loc = '' then
                        loc := H."Location Code";
                    tb.AppendLine(TableRowHtml(L."Item No.", L.Description, L.Quantity, loc));
                end;
            until L.Next() = 0;

        tb.AppendLine('</tbody></table>');
        exit(tb.ToText());
    end;

    // Arrived → ONLY posted lines on this Posted WR for THIS PO
    local procedure BuildHtmlTableFromPostedWhseRcptLines(H: Record "Posted Whse. Receipt Header"; RelatedPO: Record "Purchase Header"): Text
    var
        L: Record "Posted Whse. Receipt Line";
        tb: TextBuilder;
    begin
        tb.AppendLine(TableHeadHtml());
        L.Reset();
        L.SetRange("No.", H."No.");
        L.SetRange("Source Document", L."Source Document"::"Purchase Order");
        L.SetRange("Source No.", RelatedPO."No.");

        if L.FindSet() then
            repeat
                if L."Item No." <> '' then
                    tb.AppendLine(TableRowHtml(L."Item No.", L.Description, L.Quantity, L."Location Code"));
            until L.Next() = 0;

        tb.AppendLine('</tbody></table>');
        exit(tb.ToText());
    end;

    // ========================
    // EMAIL + HTML HELPERS
    // ========================

    local procedure SendEmail(SubjectTxt: Text; BodyHtml: Text; ToList: List of [Text])
    var
        Email: Codeunit Email;
        Msg: Codeunit "Email Message";
        Recip: List of [Text];
        Addr: Text;
    begin
        if ToList.Count() = 0 then
            exit;

        foreach Addr in ToList do
            Recip.Add(Addr);

        Msg.Create(Recip, SubjectTxt, BodyHtml, true); // HTML
        Email.Send(Msg, Enum::"Email Scenario"::Default);
    end;

    local procedure BuildRecipientListFromPO(P: Record "Purchase Header"): List of [Text]
    var
        list: List of [Text];
        primary: Text;
    begin
        primary := ResolvePOPrimaryEmail(P);
        if primary <> '' then
            list.Add(primary);
        exit(list);
    end;

    local procedure ResolvePOPrimaryEmail(P: Record "Purchase Header"): Text
    var
        Vend: Record Vendor;
        C: Record Contact;
    begin
        if P."Buy-from Contact No." <> '' then
            if C.Get(P."Buy-from Contact No.") then
                if C."E-Mail" <> '' then
                    exit(C."E-Mail");

        if Vend.Get(P."Buy-from Vendor No.") then
            if Vend."E-Mail" <> '' then
                exit(Vend."E-Mail");

        exit('');
    end;

    local procedure TableHeadHtml(): Text
    begin
        exit(
          '<table style="border-collapse:collapse;width:100%;">' +
          '<thead><tr style="background:#0b5cab;color:#fff;text-align:left;">' +
          '<th style="padding:8px;border:1px solid #ddd;">Item No.</th>' +
          '<th style="padding:8px;border:1px solid #ddd;">Description</th>' +
          '<th style="padding:8px;border:1px solid #ddd;">Qty</th>' +
          '<th style="padding:8px;border:1px solid #ddd;">Location</th>' +
          '</tr></thead><tbody>');
    end;

    local procedure TableRowHtml(ItemNo: Code[20]; Desc: Text[100]; Qty: Decimal; Loc: Code[10]): Text
    begin
        exit(
          '<tr>' +
            '<td style="padding:8px;border:1px solid #ddd;">' + HtmlEncode(Format(ItemNo)) + '</td>' +
            '<td style="padding:8px;border:1px solid #ddd;">' + HtmlEncode(Desc) + '</td>' +
            '<td style="padding:8px;border:1px solid #ddd;">' + HtmlEncode(Format(Qty)) + '</td>' +
            '<td style="padding:8px;border:1px solid #ddd;">' + HtmlEncode(Format(Loc)) + '</td>' +
          '</tr>');
    end;

    local procedure WrapHtml(Title: Text; BodyInner: Text): Text
    begin
        exit(
          '<!DOCTYPE html><html><head><meta charset="utf-8"><title>' + HtmlEncode(Title) + '</title></head>' +
          '<body style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#111;">' +
          BodyInner + '</body></html>');
    end;

    local procedure SignatureHtml(): Text
    begin
        exit(
          '<p>If you have any questions, please contact your Supply Chain Team: ' +
          '<a href="mailto:cyndy.peterson@bestwaycorp.us">cyndy.peterson@bestwaycorp.us</a> and/or ' +
          '<a href="mailto:Eric.Eichstaedt@bestwaycorp.us">Eric.Eichstaedt@bestwaycorp.us</a>.</p>');
    end;

    // Safe HTML encoding (no ConvertStr/SubstituteStr pitfalls)
    local procedure HtmlEncode(S: Text): Text
    begin
        S := ReplaceAll(S, '&', '&amp;');
        S := ReplaceAll(S, '<', '&lt;');
        S := ReplaceAll(S, '>', '&gt;');
        S := ReplaceAll(S, '"', '&quot;');
        S := ReplaceAll(S, '''', '&#39;');
        exit(S);
    end;

    local procedure ReplaceAll(TextIn: Text; FindWhat: Text; ReplaceWith: Text): Text
    var
        p: Integer;
        startIdx: Integer;
        outTxt: Text;
        rest: Text;
        chunk: Text;
    begin
        if (FindWhat = '') or (TextIn = '') then
            exit(TextIn);
        outTxt := '';
        startIdx := 1;
        p := StrPos(TextIn, FindWhat);
        while p > 0 do begin
            chunk := CopyStr(TextIn, startIdx, p - startIdx);
            outTxt += chunk + ReplaceWith;
            startIdx := p + StrLen(FindWhat);
            rest := CopyStr(TextIn, startIdx);
            p := StrPos(rest, FindWhat);
            if p > 0 then
                p := p + startIdx - 1;
        end;
        outTxt += CopyStr(TextIn, startIdx);
        exit(outTxt);
    end;

    local procedure GetETAText(P: Record "Purchase Header"): Text
    begin
        if P."Expected Receipt Date" <> 0D then
            exit(Format(P."Expected Receipt Date"));
        if P."Promised Receipt Date" <> 0D then
            exit(Format(P."Promised Receipt Date"));
        exit('');
    end;
}