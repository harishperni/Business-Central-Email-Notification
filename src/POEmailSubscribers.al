codeunit 50143 "PO Email Subscribers"
{
    // --- Shipped: fire when a Warehouse Receipt is created (no lock with Release) ---
    [EventSubscriber(ObjectType::Table, Database::"Warehouse Receipt Header", 'OnAfterInsertEvent', '', false, false)]
    local procedure OnWrInserted(var Rec: Record "Warehouse Receipt Header")
    var
        Helper: Codeunit "PO Email Helper";
        WrLine: Record "Warehouse Receipt Line";
        PH: Record "Purchase Header";
        PONos: List of [Code[20]];
        PONo: Code[20];
    begin
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

    // --- Arrived: fire when a Posted Warehouse Receipt is created (no lock with Release) ---
    [EventSubscriber(ObjectType::Table, Database::"Posted Whse. Receipt Header", 'OnAfterInsertEvent', '', false, false)]
    local procedure OnPostedWrInserted(var Rec: Record "Posted Whse. Receipt Header")
    var
        Helper: Codeunit "PO Email Helper";
        PL: Record "Posted Whse. Receipt Line";
        PH: Record "Purchase Header";
        PONos: List of [Code[20]];
        PONo: Code[20];
    begin
        PL.SetRange("No.", Rec."No.");
        PL.SetRange("Source Document", PL."Source Document"::"Purchase Order");
        if PL.FindSet() then
            repeat
                PONo := PL."Source No.";
                if (PONo <> '') and not PONos.Contains(PONo) then
                    PONos.Add(PONo);
            until PL.Next() = 0;

        foreach PONo in PONos do
            if PH.Get(PH."Document Type"::Order, PONo) then
                Helper.Notify_Arrived_OnWhseReceiptPosted(Rec, PH);
    end;
}