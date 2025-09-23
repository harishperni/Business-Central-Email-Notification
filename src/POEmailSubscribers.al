codeunit 50143 "PO Email Subscribers"
{
    var
        EmailState: Codeunit "PO Email State";

    // 1) PO Released -> "PO Created"
    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Release Purchase Document", 'OnAfterReleasePurchaseDoc', '', false, false)]
    local procedure OnAfterReleasePurchaseDoc(var PurchaseHeader: Record "Purchase Header")
    var
        Helper: Codeunit "PO Email Helper";
    begin
        if PurchaseHeader."Document Type" <> PurchaseHeader."Document Type"::Order then
            exit;
        Helper.Notify_POCreated_OnRelease(PurchaseHeader);
    end;

    // 2) WR CREATED -> "Shipped" (fires when WR header is created from PO)
    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Get Source Doc. Inbound", 'OnAfterCreateWhseReceiptHeaderFromWhseRequest', '', false, false)]
    local procedure OnAfterCreateWhseReceiptHeaderFromWhseRequest(var WhseReceiptHeader: Record "Warehouse Receipt Header")
    var
        Helper: Codeunit "PO Email Helper";
        WhseRcptLine: Record "Warehouse Receipt Line";
        PurchHeader: Record "Purchase Header";
        UniquePOs: List of [Code[20]];
        PONo: Code[20];
    begin
        WhseRcptLine.SetRange("No.", WhseReceiptHeader."No.");
        if WhseRcptLine.FindSet() then
            repeat
                if WhseRcptLine."Source Type" = Database::"Purchase Line" then begin
                    PONo := WhseRcptLine."Source No.";
                    if (PONo <> '') and not UniquePOs.Contains(PONo) then
                        UniquePOs.Add(PONo);
                end;
            until WhseRcptLine.Next() = 0;

        foreach PONo in UniquePOs do
            if PurchHeader.Get(PurchHeader."Document Type"::Order, PONo) then
                Helper.Notify_Shipped_OnWhseReceiptCreated(WhseReceiptHeader, PurchHeader);
    end;

    // 3) WR POSTED -> "Arrived"
    // Single reliable hook: posted line insert (7319). Debounce so ALL lines are visible, then send once.
    [EventSubscriber(ObjectType::Table, Database::"Posted Whse. Receipt Line", 'OnAfterInsertEvent', '', false, false)]
    local procedure Arrived_FromPostedLine(var Rec: Record "Posted Whse. Receipt Line"; RunTrigger: Boolean)
    var
        Helper: Codeunit "PO Email Helper";
        PostedHdr: Record "Posted Whse. Receipt Header";
        PwrLine: Record "Posted Whse. Receipt Line";
        PurchHeader: Record "Purchase Header";
        PONumbers: List of [Code[20]];
        PONo: Code[20];
        LineCount: Integer;
        i: Integer;
    begin
        // De-dup per posted receipt
        if not EmailState.ShouldSendArrived(Rec."No.") then begin
            // Message('DEBUG 7319: dedup blocked for %1', Rec."No.");
            exit;
        end;

        // Small debounce so *all* lines of this posted receipt are committed
        // (200ms * 8 = up to ~1.6s)
        i := 0;
        repeat
            SLEEP(200);
            i += 1;

            // Count lines for this posted receipt
            PwrLine.Reset();
            PwrLine.SetRange("No.", Rec."No.");
            PwrLine.SetRange("Source Type", Database::"Purchase Line");
            LineCount := 0;
            if PwrLine.FindSet() then
                repeat
                    LineCount += 1;
                until PwrLine.Next() = 0;
        until (LineCount > 0) or (i >= 8);

        // Message('DEBUG 7319: %1 line(s) after %2 waits for %3', LineCount, i, Rec."No.");
        if LineCount = 0 then
            exit;

        if not PostedHdr.Get(Rec."No.") then
            exit;

        // Collect distinct POs represented on this posted receipt
        PwrLine.Reset();
        PwrLine.SetRange("No.", Rec."No.");
        PwrLine.SetRange("Source Type", Database::"Purchase Line");
        if PwrLine.FindSet() then
            repeat
                PONo := PwrLine."Source No.";
                if (PONo <> '') and not PONumbers.Contains(PONo) then
                    PONumbers.Add(PONo);
            until PwrLine.Next() = 0;

        // Send one Arrived email per PO using the helper (it prints ALL posted lines for that PO)
        foreach PONo in PONumbers do
            if PurchHeader.Get(PurchHeader."Document Type"::Order, PONo) then
                Helper.Notify_Arrived_OnWhseReceiptPosted(PostedHdr, PurchHeader);

        // Mark dedup only after we actually dispatched
        EmailState.MarkArrivedSent(Rec."No.");
        // Message('DEBUG 7319: Arrived marked sent for %1', Rec."No.");
    end;
}