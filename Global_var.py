duplicate = 0
inserted = 0
expired = 0
skipped = 0
deadline_Not_given = 0
From_Date = ''
todate = ''
On_Error = 0
Total = 0
QC_Tenders = 0


def Process_End():
    print("Total: ",Total)
    print('Duplicate: ' , duplicate)
    print('Expired: ' , expired)
    print('Inserted: ' , inserted)
    print('Skipped: ' , skipped)
    print('QC Tenders: ', QC_Tenders)