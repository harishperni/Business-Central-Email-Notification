codeunit 50143 "PO Email Subscribers"
{
    // 1) PO Released -> "Created" email
    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Release Purchase Document", 'OnAfterReleasePurchaseDoc', '', false, false)]
    local procedure OnAfterReleasePurchaseDoc(var PurchaseHeader: Record "Purchase Header")
    var
        Helper: Codeunit "PO Email Helper";
    begin
        if PurchaseHeader."Document Type" <> PurchaseHeader."Document Type"::Order then
            exit;

        // Helper.Html() is a no-op now, so this is read-only and won't lock.
        Helper.Notify_POCreated_OnRelease(PurchaseHeader);
    end;

    // 2) Warehouse Receipt CREATED -> "Shipped" email
    // Use RunTrigger guard: TRUE on user-created WR, FALSE on system inserts during posting.
    [EventSubscriber(ObjectType::Table, Database::"Warehouse Receipt Header", 'OnAfterInsertEvent', '', false, false)]
    local procedure OnWhseReceiptInserted(var Rec: Record "Warehouse Receipt Header"; RunTrigger: Boolean)
    var
        Helper: Codeunit "PO Email Helper";
        WrLine: Record "Warehouse Receipt Line";
        PH: Record "Purchase Header";
        PONos: List of [Code[20]];
        PONo: Code[20];
    begin
        if not RunTrigger then
            exit; // avoid firing during posting-related inserts

        // Collect distinct POs referenced by this WR
        WrLine.SetRange("No.", Rec."No.");
        WrLine.SetRange("Source Document", WrLine."Source Document"::"Purchase Order");
        if WrLine.FindSet() then
            repeat
                PONo := WrLine."Source No.";
                if (PONo <> '') and not PONos.Contains(PONo) then
                    PONos.Add(PONo);
            until WrLine.Next() = 0;

        foreach PONo in PONos do
            if PH.Get(PH."Document Type"::Order, PONo) then
                Helper.Notify_Shipped_OnWhseReceiptCreated(Rec, PH);
    end;

    // 3) Posted Warehouse Receipt CREATED -> "Arrived" email
    [EventSubscriber(ObjectType::Table, Database::"Posted Whse. Receipt Header", 'OnAfterInsertEvent', '', false, false)]
    local procedure OnPostedWhseReceiptInserted(var Rec: Record "Posted Whse. Receipt Header"; RunTrigger: Boolean)
    var
        Helper: Codeunit "PO Email Helper";
        PwrLine: Record "Posted Whse. Receipt Line";
        PH: Record "Purchase Header";
        PONos: List of [Code[20]];
        PONo: Code[20];
    begin
        if not RunTrigger then
            exit; // ignore any non-user/system inserts

        // Collect distinct POs for this posted receipt
        PwrLine.SetRange("No.", Rec."No.");
        PwrLine.SetRange("Source Document", PwrLine."Source Document"::"Purchase Order");
        if PwrLine.FindSet() then
            repeat
                PONo := PwrLine."Source No.";
                if (PONo <> '') and not PONos.Contains(PONo) then
                    PONos.Add(PONo);
            until PwrLine.Next() = 0;

        foreach PONo in PONos do
            if PH.Get(PH."Document Type"::Order, PONo) then
                Helper.Notify_Arrived_OnWhseReceiptPosted(Rec, PH);
    end;
}