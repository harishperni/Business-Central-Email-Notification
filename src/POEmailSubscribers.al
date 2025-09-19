codeunit 50143 "PO Email Subscribers"
{
    // 1) PO Released -> "Created"
    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Release Purchase Document", 'OnAfterReleasePurchaseDoc', '', false, false)]
    local procedure OnAfterReleasePurchaseDoc(var PurchaseHeader: Record "Purchase Header")
    var
        Helper: Codeunit "PO Email Helper";
    begin
        if PurchaseHeader."Document Type" = PurchaseHeader."Document Type"::Order then
            Helper.Notify_POCreated_OnRelease(PurchaseHeader);
    end;

    // 2) Warehouse Receipt CREATED -> "Shipped"
    [EventSubscriber(ObjectType::Table, Database::"Warehouse Receipt Header", 'OnAfterInsertEvent', '', false, false)]
    local procedure OnWhseReceiptInserted(var Rec: Record "Warehouse Receipt Header")
    var
        Helper: Codeunit "PO Email Helper";
        L: Record "Warehouse Receipt Line";
        PurchHeader: Record "Purchase Header";
        PONumbers: List of [Code[20]];
        PONo: Code[20];
    begin
        // Collect unique POs on this receipt
        L.SetRange("No.", Rec."No.");
        L.SetRange("Source Document", L."Source Document"::"Purchase Order");
        if L.FindSet() then
            repeat
                PONo := L."Source No.";
                if (PONo <> '') and not PONumbers.Contains(PONo) then
                    PONumbers.Add(PONo);
            until L.Next() = 0;

        // Notify once per PO represented on the WR
        foreach PONo in PONumbers do
            if PurchHeader.Get(PurchHeader."Document Type"::Order, PONo) then
                Helper.Notify_Shipped_OnWhseReceiptCreated(Rec, PurchHeader);
    end;

    // 3) Posted Warehouse Receipt CREATED -> "Arrived"
    [EventSubscriber(ObjectType::Table, Database::"Posted Whse. Receipt Header", 'OnAfterInsertEvent', '', false, false)]
    local procedure OnPostedWhseReceiptInserted(var Rec: Record "Posted Whse. Receipt Header")
    var
        Helper: Codeunit "PO Email Helper";
        L: Record "Posted Whse. Receipt Line";
        PurchHeader: Record "Purchase Header";
        PONumbers: List of [Code[20]];
        PONo: Code[20];
    begin
        // Collect unique POs on this posted receipt
        L.SetRange("No.", Rec."No.");
        L.SetRange("Source Document", L."Source Document"::"Purchase Order");
        if L.FindSet() then
            repeat
                PONo := L."Source No.";
                if (PONo <> '') and not PONumbers.Contains(PONo) then
                    PONumbers.Add(PONo);
            until L.Next() = 0;

        // Notify once per PO represented on the posted WR
        foreach PONo in PONumbers do
            if PurchHeader.Get(PurchHeader."Document Type"::Order, PONo) then
                Helper.Notify_Arrived_FromPostedReceipt(Rec, PurchHeader);
    end;
}