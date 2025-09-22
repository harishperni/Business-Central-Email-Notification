codeunit 50143 "PO Email Subscribers"
{
    // Wires Business Central events to your helper.
    // Uses your existing SingleInstance codeunit 50146 "PO Email State" for de-dup.

    var
        EmailState: Codeunit "PO Email State";

    // ---------------------------------------------------------
    // 1) PO Released  -> "PO Created"
    // ---------------------------------------------------------
    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Release Purchase Document", 'OnAfterReleasePurchaseDoc', '', false, false)]
    local procedure OnAfterReleasePurchaseDoc(var PurchaseHeader: Record "Purchase Header")
    var
        Helper: Codeunit "PO Email Helper";
    begin
        if PurchaseHeader."Document Type" <> PurchaseHeader."Document Type"::Order then
            exit;

        // Telemetry
        Session.LogMessage(
            'po.release',
            StrSubstNo('PO released: %1', PurchaseHeader."No."),
            Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
            'PONo', PurchaseHeader."No.");

        Helper.Notify_POCreated_OnRelease(PurchaseHeader);
    end;

    // ---------------------------------------------------------
    // 2) WR CREATED -> "Shipped"
    //    Use 5751 creation event so it fires when the WR header is actually created.
    // ---------------------------------------------------------
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
            Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
            'WR', WhseReceiptHeader."No.");

        // Collect distinct POs on this WR and send one email per PO
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

    // ---------------------------------------------------------
    // 3) WR POSTED -> "Arrived"
    //    Listen to multiple reliable points (5760/7318/7319) and de-dup via EmailState.
    //    Your environment shows Posted Header = 7318, Posted Line = 7319.
    // ---------------------------------------------------------

    // 3A) Preferred: codeunit 5760 "Whse.-Post Receipt"
    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Whse.-Post Receipt", 'OnAfterPostedWhseRcptHeaderInsert', '', false, false)]
    local procedure Arrived_FromPostCodeunit(var PostedWhseReceiptHeader: Record "Posted Whse. Receipt Header"; WarehouseReceiptHeader: Record "Warehouse Receipt Header")
    begin
        DispatchArrived(PostedWhseReceiptHeader."No.", WarehouseReceiptHeader."No.", '5760');
    end;

    // 3B) Fallback: table insert of Posted Header (7318)
    [EventSubscriber(ObjectType::Table, Database::"Posted Whse. Receipt Header", 'OnAfterInsertEvent', '', false, false)]
    local procedure Arrived_FromPostedHeader(var Rec: Record "Posted Whse. Receipt Header"; RunTrigger: Boolean)
    begin
        DispatchArrived(Rec."No.", '', '7318');
    end;

    // 3C) Last resort: first Posted Line (7319) drives dispatch
    [EventSubscriber(ObjectType::Table, Database::"Posted Whse. Receipt Line", 'OnAfterInsertEvent', '', false, false)]
    local procedure Arrived_FromPostedLine(var Rec: Record "Posted Whse. Receipt Line"; RunTrigger: Boolean)
    begin
        // Only act once per posted receipt
        if not EmailState.ShouldSendArrived(Rec."No.") then
            exit;
        DispatchArrived(Rec."No.", '', '7319');
    end;

    // -----------------------
    // Shared Arrived dispatcher
    // -----------------------
    local procedure DispatchArrived(PostedNo: Code[20]; FromUnpostedNo: Code[20]; SourceTag: Text[10])
    var
        Helper: Codeunit "PO Email Helper";
        PostedHdr: Record "Posted Whse. Receipt Header";
        PostedLine: Record "Posted Whse. Receipt Line";
        PurchHeader: Record "Purchase Header";
        PONumbers: List of [Code[20]];
        PONo: Code[20];
        HadLines: Boolean;
    begin
        if PostedNo = '' then
            exit;

        // De-dup (uses your existing state codeunit)
        if not EmailState.ShouldSendArrived(PostedNo) then
            exit;

        Session.LogMessage(
            'wr.arrived.dispatch',
            StrSubstNo('Arrived dispatch for Posted WR %1 (from:%2, src:%3)', PostedNo, FromUnpostedNo, SourceTag),
            Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
            'PostedWR', PostedNo);

        if not PostedHdr.Get(PostedNo) then
            exit;

        // Collect PO(s) present on this *posted* receipt (what actually arrived)
        PostedLine.SetRange("No.", PostedNo);
        if PostedLine.FindSet() then
            repeat
                HadLines := true;
                if PostedLine."Source Type" = Database::"Purchase Line" then begin
                    PONo := PostedLine."Source No.";
                    if (PONo <> '') and not PONumbers.Contains(PONo) then
                        PONumbers.Add(PONo);
                end;
            until PostedLine.Next() = 0;

        if not HadLines then begin
            // Nothing posted? Log and bail.
            Session.LogMessage(
                'wr.arrived.nolines',
                StrSubstNo('No posted lines found for %1', PostedNo),
                Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher,
                'PostedWR', PostedNo);
            exit;
        end;

        foreach PONo in PONumbers do
            if PurchHeader.Get(PurchHeader."Document Type"::Order, PONo) then
                Helper.Notify_Arrived_OnWhseReceiptPosted(PostedHdr, PurchHeader);

        EmailState.MarkArrivedSent(PostedNo);
    end;
}