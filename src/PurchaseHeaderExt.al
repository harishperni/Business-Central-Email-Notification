tableextension 50148 "PO Notif Ext" extends "Purchase Header"
{
    fields
    {
        field(50990; "Notification Emails"; Text[250])
        {
            DataClassification = CustomerContent;
            Caption = 'Notification Emails';
            ToolTip = 'Extra recipients (comma or semicolon separated).';
        }
    }
}
