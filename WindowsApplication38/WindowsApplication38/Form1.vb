Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.RowCount = 3
        DataGridView1.Rows(0).Cells(0).Value = "شيبس"
        DataGridView1.Rows(0).Cells(1).Value = "10"
        DataGridView1.Rows(0).Cells(2).Value = "15"
        DataGridView1.Rows(1).Cells(0).Value = "شوكولاته"
        DataGridView1.Rows(1).Cells(1).Value = "5"
        DataGridView1.Rows(1).Cells(2).Value = "10"
        DataGridView1.Rows(2).Cells(0).Value = "بيبسي"
        DataGridView1.Rows(2).Cells(1).Value = "4"
        DataGridView1.Rows(2).Cells(2).Value = "6"
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Ex = CreateObject("Excel.Application") 'استدعاء برنامج الاكسل
        Ex.visible = True
        'Ex.Visible = True
        Ex.Workbooks.Add() 'انشاء ورقة عمل
        Ex.Cells.Range("A1:D1").Merge()  ' دمج الخلايا
        Ex.Cells(1, 1).Interior.ColorIndex = 15 ' تغيير اللون الداخلي للخليه
        Ex.Cells.Range("B3:D3").Merge() ''دمج الخلايا
        Ex.Cells.Range("B2:D2").Merge() ''دمج الخلايا
        Ex.Cells(1, 1).Value = "فاتورة مبيعات" ''كتابة نصوص
        Ex.Cells(2, 1).Value = "التاريخ" ''كتابة نصوص
        Ex.Cells(3, 1).Value = "رقم الفاتورة" ''كتابة نصوص
        Ex.Cells(2, 2).Value = Date.Now.ToString("dd/MM/yyyy") ''كتابة تاريخ اليوم
        Ex.Cells(3, 2).Value = TextBox1.Text ''كتابة نصوص
        Ex.Columns.HorizontalAlignment = 3 '   1 من اليمين الى اليسار ورقم 2 من اليسار الى اليمين ورقم 3 في الوسط
        ''عرض العامود
        'Ex.Cells(5).ColumnWidth = 12
        ''========================
        For i As Integer = 0 To DataGridView1.Columns.Count - 1
            For j As Integer = 0 To DataGridView1.Rows.Count - 1

                Ex.Cells(6, i + 1).Value = DataGridView1.Columns(i).HeaderText 'جلب اسماء الاعمدة للجدول
                Ex.Cells(j + 7, i + 1).Value = DataGridView1.Rows(j).Cells(i).Value 'جلب البيانات من الداتا كريد فيو
                Ex.Cells(j + 7, i + 1).Borders.Weight = 2 ' اضافة اطار جدول
                Ex.Cells(j + 6, i + 1).Borders.Weight = 2 ' اضافة اطار جدول
            Next
        Next
        '================================================
        'Ex.Cells(j + 1).ColumnWidth = 30 'التحكم في عرض العمود
        'Ex.Columns.Font.Name = "Times New Roman" 'اسم الخط المستخدم في الجدول
        'Ex.Rows.Item(j + 6).Font.Bold = 1 'تحويل الخط الى غامق
        'Ex.Rows.Item(j + 6).Font.size = 11 'حجم الخط
        'Ex.Columns("b:S").EntireColumn.AutoFit() 'جعل الحقول ملائمة لحجم الكتابة
    End Sub
End Class
