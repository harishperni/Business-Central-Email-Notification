codeunit 50142 "PO Email Helper"
{
    SingleInstance = true;

    var
        Email: Codeunit Email;
        EmailMsg: Codeunit "Email Message";

    // =========================================================
    // PO CREATED  (from PO lines)
    // =========================================================
    procedure Notify_POCreated_OnRelease(RelatedPO: Record "Purchase Header")
    var
        subj: Text;
        body: TextBuilder;
        html: Text;
    begin
        subj := StrSubstNo('PO %1 Created', RelatedPO."No.");

        body.AppendLine('<p>Dear Colleague,</p>');
        body.AppendLine(StrSubstNo(
            '<p>PO <strong>#%1</strong> has been <strong>Released</strong>. The items and quantities are:</p>',
            Html(RelatedPO."No.")));

        body.AppendLine(BuildHtmlTableFromPOLines(RelatedPO));
        body.AppendLine(BuildFooter());

        html := WrapHtml(subj, body.ToText());

        EmailMsg.Create(BuildRecipientListFromPO(RelatedPO), subj, html, true);
        Email.Send(EmailMsg, Enum::"Email Scenario"::Default);
    end;

    // =========================================================
    // SHIPPED  (Warehouse Receipt Created)
    // =========================================================
    procedure Notify_Shipped_OnWhseReceiptCreated(WhseRcptHeader: Record "Warehouse Receipt Header"; RelatedPO: Record "Purchase Header")
    var
        subj: Text;
        body: TextBuilder;
        html: Text;
        etaDate: Date;
        etaTxt: Text;
    begin
        // Resolve ETA: header Expected, then Promised, else earliest line date
        etaDate := GetPOETA(RelatedPO);
        if etaDate <> 0D then
            etaTxt := Format(etaDate);

        subj := StrSubstNo('PO %1 Shipped', RelatedPO."No.");

        body.AppendLine('<p>Dear Colleague,</p>');
        if etaTxt <> '' then
            body.AppendLine(StrSubstNo(
                '<p>PO <strong>#%1</strong> has <strong>Shipped!!</strong> The estimated arrival date is <strong>%2</strong>. The items and quantities are:</p>',
                Html(RelatedPO."No."), Html(etaTxt)))
        else
            body.AppendLine(StrSubstNo(
                '<p>PO <strong>#%1</strong> has <strong>Shipped!!</strong>. The items and quantities are:</p>',
                Html(RelatedPO."No.")));

        body.AppendLine(BuildHtmlTableFromWhseRcptLines(WhseRcptHeader, RelatedPO."No."));
        body.AppendLine(BuildFooter());

        html := WrapHtml(subj, body.ToText());

        EmailMsg.Create(BuildRecipientListFromPO(RelatedPO), subj, html, true);
        Email.Send(EmailMsg, Enum::"Email Scenario"::Default);

        Session.LogMessage(
            'wr.shipped.sent',
            StrSubstNo('Shipped email sent. WR=%1 PO=%2 ETA=%3', WhseRcptHeader."No.", RelatedPO."No.", etaTxt),
            Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
            'WR', WhseRcptHeader."No.");
    end;

    // =========================================================
    // ARRIVED  (Posted Warehouse Receipt Created)
    //  - Resolves the MOST-RECENT Posted WR that has lines for this PO
    //  - Renders ALL posted lines for that Posted WR & PO
    // =========================================================
    procedure Notify_Arrived_OnWhseReceiptPosted(PostedHdr: Record "Posted Whse. Receipt Header"; RelatedPO: Record "Purchase Header")
    var
        subj: Text;
        body: TextBuilder;
        html: Text;
        UsePostedHdr: Record "Posted Whse. Receipt Header";
        PwrLine: Record "Posted Whse. Receipt Line";
        latestNo: Code[20];
        latestDate: Date;
        cnt: Integer;
    begin
        subj := StrSubstNo('PO %1 Arrived', RelatedPO."No.");

        // --- Find the most-recent Posted WR that actually has lines for this PO
        //     (Sometimes the event passes a header earlier/later than we want.)
        PwrLine.Reset();
        PwrLine.SetRange("Source Type", Database::"Purchase Line");
        PwrLine.SetRange("Source No.", RelatedPO."No.");
        PwrLine.SetCurrentKey("Posting Date", "No.");
        if PwrLine.FindSet() then
            repeat
                if (PwrLine."Posting Date" > latestDate) or
                   ((PwrLine."Posting Date" = latestDate) and (PwrLine."No." > latestNo)) then begin
                    latestDate := PwrLine."Posting Date";
                    latestNo := PwrLine."No.";
                end;
            until PwrLine.Next() = 0;

        if (latestNo <> '') and UsePostedHdr.Get(latestNo) then
            Session.LogMessage(
                'wr.arrived.latest',
                StrSubstNo('Resolved latest Posted WR for PO %1 -> %2 (%3)',
                    RelatedPO."No.", latestNo, Format(latestDate)),
                Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
                'PostedWR', latestNo)
        else begin
            // fallback to the header provided by the event
            UsePostedHdr := PostedHdr;
            Session.LogMessage(
                'wr.arrived.fallback',
                StrSubstNo('Using event Posted WR %1 for PO %2', PostedHdr."No.", RelatedPO."No."),
                Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
                'PostedWR', PostedHdr."No.");
        end;

        // Count rows for the table (debug/telemetry)
        PwrLine.Reset();
        PwrLine.SetRange("No.", UsePostedHdr."No.");
        PwrLine.SetRange("Source Type", Database::"Purchase Line");
        PwrLine.SetRange("Source No.", RelatedPO."No.");
        if PwrLine.FindSet() then
            repeat
                if PwrLine."Item No." <> '' then
                    cnt += 1;
            until PwrLine.Next() = 0;

        Session.LogMessage(
            'wr.arrived.count',
            StrSubstNo('Arrived rows=%1 for PostedWR=%2 / PO=%3', cnt, UsePostedHdr."No.", RelatedPO."No."),
            Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
            'PostedWR', UsePostedHdr."No.");

        // Compose email
        body.AppendLine('<p>Dear Colleague,</p>');
        body.AppendLine(StrSubstNo(
            '<p>PO <strong>#%1</strong> has <strong>Arrived</strong> at our Bestway USA Chandler Warehouse!! ' +
            'Please allow 3–5 days for Container Unloading &amp; Putaways. The posted items and quantities are:</p>',
            Html(RelatedPO."No.")));

        // Full table of ALL posted lines for the (resolved) posted receipt & this PO
        body.AppendLine(BuildHtmlTableFromPostedWRLines_All(UsePostedHdr, RelatedPO."No."));
        body.AppendLine(BuildFooter());

        html := WrapHtml(subj, body.ToText());

        EmailMsg.Create(BuildRecipientListFromPO(RelatedPO), subj, html, true);
        Email.Send(EmailMsg, Enum::"Email Scenario"::Default);

        Session.LogMessage(
            'wr.arrived.sent',
            StrSubstNo('Arrived email sent. PostedWR=%1 PO=%2 Lines=%3', UsePostedHdr."No.", RelatedPO."No.", cnt),
            Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
            'PostedWR', UsePostedHdr."No.");
    end;

    // =========================================================
    // Helpers (HTML + Recipients + Tables + ETA)
    // =========================================================

    local procedure Html(t: Text): Text
    begin
        // Keep simple to avoid ALConvertStr issues
        exit(t);
    end;

    local procedure WrapHtml(subj: Text; body: Text): Text
    begin
        exit(
          '<html><head><meta charset="utf-8">' +
          '<style>' +
          'table{border-collapse:collapse;margin:14px 0;width:100%;max-width:900px;}' +
          'th,td{border:1px solid #d9d9d9;padding:8px 10px;}' +
          'th{background:#005fb8;color:#fff;text-align:left;}' +
          'td.num{ text-align:right; }' +
          '</style></head><body>' +
          '<h2>' + subj + '</h2>' +
          body +
          '</body></html>'
        );
    end;

    local procedure BuildFooter(): Text
    begin
        exit(
          '<p>If you have any questions, please contact your Supply Chain Team: ' +
          '<a href="mailto:cyndy.peterson@bestwaycorp.us">cyndy.peterson@bestwaycorp.us</a> ' +
          'and/or ' +
          '<a href="mailto:Eric.Eichstaedt@bestwaycorp.us">Eric.Eichstaedt@bestwaycorp.us</a>.</p>'
        );
    end;

    local procedure BuildRecipientListFromPO(var PurchHeader: Record "Purchase Header"): List of [Text]
    var
        list: List of [Text];
        Vendor: Record Vendor;
        Contact: Record Contact;
        emailTxt: Text;
    begin
        // Prefer Buy-from Contact email; else Vendor email
        if (PurchHeader."Buy-from Contact No." <> '') and Contact.Get(PurchHeader."Buy-from Contact No.") then
            if Contact."E-Mail" <> '' then
                emailTxt := Contact."E-Mail";

        if (emailTxt = '') and Vendor.Get(PurchHeader."Buy-from Vendor No.") then
            if Vendor."E-Mail" <> '' then
                emailTxt := Vendor."E-Mail";

        if emailTxt <> '' then
            list.Add(emailTxt);

        exit(list);
    end;

    // ---------- Tables ----------

    local procedure BuildHtmlTableFromPOLines(var PurchHeader: Record "Purchase Header"): Text
    var
        L: Record "Purchase Line";
        sb: TextBuilder;
    begin
        sb.AppendLine('<table>');
        sb.AppendLine('<tr><th>Item No.</th><th>Description</th><th style="text-align:right;">Qty</th><th>Location</th></tr>');

        L.SetRange("Document Type", PurchHeader."Document Type");
        L.SetRange("Document No.", PurchHeader."No.");
        if L.FindSet() then
            repeat
                if (L.Type = L.Type::Item) and (L."No." <> '') then
                    sb.AppendLine(
                      '<tr>' +
                      StrSubstNo('<td>%1</td>', Html(L."No.")) +
                      StrSubstNo('<td>%1</td>', Html(L.Description)) +
                      StrSubstNo('<td class="num">%1</td>', Format(L.Quantity)) +
                      StrSubstNo('<td>%1</td>', Html(L."Location Code")) +
                      '</tr>'
                    );
            until L.Next() = 0;

        sb.AppendLine('</table>');
        exit(sb.ToText());
    end;

    local procedure BuildHtmlTableFromWhseRcptLines(var WhseRcptHeader: Record "Warehouse Receipt Header"; PONo: Code[20]): Text
    var
        L: Record "Warehouse Receipt Line";
        sb: TextBuilder;
    begin
        sb.AppendLine('<table>');
        sb.AppendLine('<tr><th>Item No.</th><th>Description</th><th style="text-align:right;">Qty</th><th>Location</th></tr>');

        L.SetRange("No.", WhseRcptHeader."No.");
        L.SetRange("Source Type", Database::"Purchase Line");
        L.SetRange("Source No.", PONo);
        if L.FindSet() then
            repeat
                if L."Item No." <> '' then
                    sb.AppendLine(
                      '<tr>' +
                      StrSubstNo('<td>%1</td>', Html(L."Item No.")) +
                      StrSubstNo('<td>%1</td>', Html(L.Description)) +
                      StrSubstNo('<td class="num">%1</td>', Format(L.Quantity)) +
                      StrSubstNo('<td>%1</td>', Html(WhseRcptHeader."Location Code")) +
                      '</tr>'
                    );
            until L.Next() = 0;

        sb.AppendLine('</table>');
        exit(sb.ToText());
    end;

    // Renamed builder (avoid signature conflicts); uses Quantity only
    local procedure BuildHtmlTableFromPostedWRLines_All(PostedHdr: Record "Posted Whse. Receipt Header"; PONo: Code[20]): Text
    var
        L: Record "Posted Whse. Receipt Line";
        sb: TextBuilder;
        hadRows: Boolean;
    begin
        sb.AppendLine('<table>');
        sb.AppendLine('<tr><th>Item No.</th><th>Description</th><th style="text-align:right;">Qty</th><th>Location</th></tr>');

        // Only posted lines for THIS posted receipt & THIS PO
        L.SetRange("No.", PostedHdr."No.");
        L.SetRange("Source Type", Database::"Purchase Line");
        L.SetRange("Source No.", PONo);

        if L.FindSet() then
            repeat
                if L."Item No." <> '' then begin
                    hadRows := true;
                    sb.AppendLine(
                      '<tr>' +
                      StrSubstNo('<td>%1</td>', Html(L."Item No.")) +
                      StrSubstNo('<td>%1</td>', Html(L.Description)) +
                      StrSubstNo('<td class="num">%1</td>', Format(L.Quantity)) +
                      StrSubstNo('<td>%1</td>', Html(PostedHdr."Location Code")) +
                      '</tr>'
                    );
                end;
            until L.Next() = 0
        else
            sb.AppendLine('<tr><td colspan="4" style="padding:10px;border:1px solid #d9d9d9;">No posted lines found for this PO on this posted receipt.</td></tr>');

        sb.AppendLine('</table>');
        exit(sb.ToText());
    end;

    // ---------- ETA helper ----------

    local procedure GetPOETA(var PO: Record "Purchase Header"): Date
    var
        PL: Record "Purchase Line";
        best: Date;
        d: Date;
    begin
        // Header dates first
        if PO."Expected Receipt Date" <> 0D then
            exit(PO."Expected Receipt Date");
        if PO."Promised Receipt Date" <> 0D then
            exit(PO."Promised Receipt Date");

        // Else earliest non-blank date from lines (Expected, else Promised)
        best := DMY2DATE(31, 12, 9999);
        PL.SetRange("Document Type", PO."Document Type");
        PL.SetRange("Document No.", PO."No.");
        PL.SetRange(Type, PL.Type::Item);
        if PL.FindSet() then
            repeat
                d := 0D;
                if PL."Expected Receipt Date" <> 0D then
                    d := PL."Expected Receipt Date"
                else
                    if PL."Promised Receipt Date" <> 0D then
                        d := PL."Promised Receipt Date";

                if (d <> 0D) and (d < best) then
                    best := d;
            until PL.Next() = 0;

        if best <> DMY2DATE(31, 12, 9999) then
            exit(best);

        exit(0D);
    end;
}