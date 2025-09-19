pageextension 50147 "PO Notif Ext Pg" extends "Purchase Order"
{
    layout
    {
        addlast(General)
        {
            field("Notification Emails"; Rec."Notification Emails")
            {
                ApplicationArea = All;
            }
        }
    }
}
