codeunit 50143 "PO Email Subscribers"
{
    // 1) PO Released -> "PO Created"
    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Release Purchase Document", 'OnAfterReleasePurchaseDoc', '', false, false)]
    local procedure OnAfterReleasePurchaseDoc(var PurchaseHeader: Record "Purchase Header")
    var
        Helper: Codeunit "PO Email Helper";
    begin
        if PurchaseHeader."Document Type" <> PurchaseHeader."Document Type"::Order then
            exit;

        Session.LogMessage(
            'po.release',
            StrSubstNo('PO released: %1', PurchaseHeader."No."),
            Verbosity::Normal,
            DataClassification::SystemMetadata,
            TelemetryScope::ExtensionPublisher,
            'PONo', PurchaseHeader."No.");

        Helper.Notify_POCreated_OnRelease(PurchaseHeader);
    end;

    // 2) WR CREATED (reliable path) -> "Shipped"
    //    Use Get Source Doc. Inbound event that fires after header creation from PO
    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Get Source Doc. Inbound", 'OnAfterCreateWhseReceiptHeaderFromWhseRequest', '', false, false)]
    local procedure OnAfterCreateWhseReceiptHeaderFromWhseRequest(var WhseReceiptHeader: Record "Warehouse Receipt Header")
    var
        Helper: Codeunit "PO Email Helper";
        WhseRcptLine: Record "Warehouse Receipt Line";
        PurchHeader: Record "Purchase Header";
        UniquePOs: List of [Code[20]];
        PONo: Code[20];
    begin
        Session.LogMessage(
            'wr.create',
            StrSubstNo('WR created: %1', WhseReceiptHeader."No."),
            Verbosity::Normal,
            DataClassification::SystemMetadata,
            TelemetryScope::ExtensionPublisher,
            'WR', WhseReceiptHeader."No.");

        // Collect related POs from this receipt’s lines (so the email shows only its items)
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

    // ARRIVED v2 — fire from Codeunit 5760 after posted header is created
    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Whse.-Post Receipt", 'OnAfterPostedWhseRcptHeaderInsert', '', false, false)]
    local procedure OnAfterPostedWhseRcptHeaderInsert(var PostedWhseReceiptHeader: Record "Posted Whse. Receipt Header"; WarehouseReceiptHeader: Record "Warehouse Receipt Header")
    var
        Helper: Codeunit "PO Email Helper";
        PostedLine: Record "Posted Whse. Receipt Line";
        PurchHeader: Record "Purchase Header";
        UniquePOs: List of [Code[20]];
        PONo: Code[20];
    begin
        // telemetry (correct overload requires a dimension/value)
        Session.LogMessage(
            'wr.post.codeunit',
            StrSubstNo('Posted WR: %1 (from %2)', PostedWhseReceiptHeader."No.", WarehouseReceiptHeader."No."),
            Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
            'PostedWR', PostedWhseReceiptHeader."No.");

        // Collect POs from POSTED lines (what actually arrived)
        PostedLine.SetRange("No.", PostedWhseReceiptHeader."No.");
        if PostedLine.FindSet() then
            repeat
                if PostedLine."Source Type" = Database::"Purchase Line" then begin
                    PONo := PostedLine."Source No.";
                    if (PONo <> '') and not UniquePOs.Contains(PONo) then
                        UniquePOs.Add(PONo);
                end;
            until PostedLine.Next() = 0;

        foreach PONo in UniquePOs do
            if PurchHeader.Get(PurchHeader."Document Type"::Order, PONo) then
                Helper.Notify_Arrived_OnWhseReceiptPosted(PostedWhseReceiptHeader, PurchHeader);
    end;

    // ARRIVED v2 fallback — also listen to the table insert (covers alternate posting paths)
    [EventSubscriber(ObjectType::Table, Database::"Posted Whse. Receipt Header", 'OnAfterInsertEvent', '', false, false)]
    local procedure OnPostedWhseReceiptInserted_Table(var Rec: Record "Posted Whse. Receipt Header"; RunTrigger: Boolean)
    var
        Helper: Codeunit "PO Email Helper";
        PostedLine: Record "Posted Whse. Receipt Line";
        PurchHeader: Record "Purchase Header";
        UniquePOs: List of [Code[20]];
        PONo: Code[20];
    begin
        Session.LogMessage(
            'wr.post.table',
            StrSubstNo('Posted WR inserted (table): %1 RunTrigger=%2', Rec."No.", Format(RunTrigger)),
            Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
            'PostedWR', Rec."No.");

        // Collect POs from POSTED lines
        PostedLine.SetRange("No.", Rec."No.");
        if PostedLine.FindSet() then
            repeat
                if PostedLine."Source Type" = Database::"Purchase Line" then begin
                    PONo := PostedLine."Source No.";
                    if (PONo <> '') and not UniquePOs.Contains(PONo) then
                        UniquePOs.Add(PONo);
                end;
            until PostedLine.Next() = 0;

        foreach PONo in UniquePOs do
            if PurchHeader.Get(PurchHeader."Document Type"::Order, PONo) then
                Helper.Notify_Arrived_OnWhseReceiptPosted(Rec, PurchHeader);
    end;
}
