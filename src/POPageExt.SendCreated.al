pageextension 50147 "PO Send Created Ext" extends "Purchase Order"
{
    actions
    {
        addlast(Processing)
        {
            action(SendPOCreatedEmail)
            {
                ApplicationArea = All;
                Caption = 'Send Created Email';
                Image = Email;
                ToolTip = 'Send the PO Created email (runs outside the Release transaction).';

                trigger OnAction()
                var
                    Helper: Codeunit "PO Email Helper";
                    PH: Record "Purchase Header";
                begin
                    PH.Get(Rec."Document Type", Rec."No.");
                    Helper.Notify_POCreated_OnRelease(PH);
                    Message('Created email sent for PO %1.', PH."No.");
                end;
            }
        }
    }
}