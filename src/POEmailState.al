codeunit 50146 "PO Email State"
{
    SingleInstance = true;

    var
        ShippedSent: Dictionary of [Code[20], Boolean];
        ArrivedSent: Dictionary of [Code[20], Boolean];

    procedure ShouldSendShipped(ReceiptNo: Code[20]): Boolean
    begin
        exit(not ShippedSent.ContainsKey(ReceiptNo));
    end;

    procedure MarkShippedSent(ReceiptNo: Code[20])
    begin
        ShippedSent.Set(ReceiptNo, true);
    end;

    procedure ShouldSendArrived(PostedReceiptNo: Code[20]): Boolean
    begin
        exit(not ArrivedSent.ContainsKey(PostedReceiptNo));
    end;

    procedure MarkArrivedSent(PostedReceiptNo: Code[20])
    begin
        ArrivedSent.Set(PostedReceiptNo, true);
    end;

    procedure ResetAll()
    begin
        CLEAR(ShippedSent);
        CLEAR(ArrivedSent);
    end;
}