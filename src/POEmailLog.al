table 50145 "PO Email Log"
{
    DataClassification = ToBeClassified;

    fields
    {
        field(1; "Kind"; Code[10])      // must stay Code[10] to match existing data
        {
            Caption = 'Kind';
        }
        field(2; "Receipt No."; Code[20])
        {
            Caption = 'Receipt No.';
        }
        field(3; "PO No."; Code[20])
        {
            Caption = 'PO No.';
        }
        field(4; "Inserted At"; DateTime)
        {
            Caption = 'Inserted At';
        }
    }

    keys
    {
        key(PK; "Kind", "Receipt No.", "PO No.")
        {
            Clustered = true;
        }
    }
}