using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Xml.Linq;


public class SinhVien    //Lop SV
{
    //Cac thong tin ve SV
    public string MaSV { get; set; }
    public string HoTen { get; set; }
    public int NamSinh { get; set; }
    public double DiemTT { get; set; }
    public double DiemTB { get; set; }

    public SinhVien(string maSV, string hoTen, int namSinh, double diemTT, double diemTB)
    {
        MaSV = maSV;
        HoTen = hoTen;
        NamSinh = namSinh;
        DiemTT = diemTT;
        DiemTB = diemTB;
    }


}


public class HashTable
{
    private int SIZE = 100;
    private List<SinhVien>[] table;

    public HashTable()
    {
        table = new List<SinhVien>[SIZE];
        for (int i = 0; i < SIZE; i++)
        {
            table[i] = new List<SinhVien>();
        }
    }

    private int HashFunction(string key)
    {
        return Convert.ToInt32(key) % SIZE;
    }

    public void Add1(SinhVien sv)    //Phương thức add cho người dùng (đã kiểm tra trùng bằng Find() từ trước
    {
        int index = HashFunction(sv.MaSV);
        table[index].Add(sv);
        Console.WriteLine("Thêm sinh viên mới thành công!");
    }

    public void Add(SinhVien sv)        //Phương thức add sinh viên trong danh sách
    {
        int index = HashFunction(sv.MaSV);
        if (!table[index].Contains(sv))
        {
            table[index].Add(sv);
        }
    }

    public SinhVien Find(string maSV)
    {
        int index = HashFunction(maSV);
        if (table[index].Count > 0)
        {
            for (int i = 0; i < table[index].Count; i++)
            {
                if (string.Compare(maSV, table[index][i].MaSV) == 0)
                {
                    return table[index][i];
                }
            }
        }
        return null;
    }

    public void Delete(string maSV)
    {
        int index = HashFunction(maSV);
        if (table[index].Count > 0)
        {
            for (int i = table[index].Count - 1; i >= 0; i--)
            {
                if (string.Compare(maSV, table[index][i].MaSV) == 0)
                {
                    table[index].RemoveAt(i);
                    Console.WriteLine("Xóa thành công sinh viên với mã " + maSV);
                    return;
                }
            }
        }
        Console.WriteLine($"Sinh viên có mã {maSV} không tồn tại!");
    }

    public void UpdateMark(string maSV, double mark)   //Phương thức của người dùng chắc chắn maSV tồn tại (đã kiểm tra = Find())
    {
        int index = HashFunction(maSV);
        for (int i = 0; i < table[index].Count; i++)
        {
            if (string.Compare(maSV, table[index][i].MaSV) == 0)
            {
                table[index][i].DiemTB = mark;
                Console.WriteLine($"Cập nhật điểm cho sinh viên có mã {maSV} thành công!");
                return;
            }
        }
    }

    public void ShowAllStudent()
    {
        foreach (List<SinhVien> list in table)
        {
            foreach (SinhVien sv in list)
            {
                Console.WriteLine("Mã sinh viên: " + sv.MaSV);
                Console.WriteLine("Họ tên: " + sv.HoTen);
                Console.WriteLine("Năm sinh: " + sv.NamSinh);
                Console.WriteLine("Điểm trúng tuyển: " + sv.DiemTT);
                Console.WriteLine("Điểm trung bình: " + sv.DiemTB);
                Console.WriteLine();

            }
        }
    }


    public bool ExportFile(string sheetName)
    {
        using (ExcelPackage package = new ExcelPackage(new FileInfo(Path.Combine(Directory.GetCurrentDirectory(), "QLSinhVien.xlsx"))))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null)
            {
                worksheet = package.Workbook.Worksheets.Add(sheetName);
                // Tiêu đề cột
                worksheet.Cells[1, 1].Value = "MSSV";
                worksheet.Cells[1, 2].Value = "Họ tên";
                worksheet.Cells[1, 3].Value = "Năm sinh";
                worksheet.Cells[1, 4].Value = "Điểm trúng tuyển";
                worksheet.Cells[1, 5].Value = "Điểm trung bình";

                int row = 2; // Bắt đầu từ hàng 2 để viết dữ liệu

                foreach (List<SinhVien> list in table)
                {
                    foreach (SinhVien sv in list)
                    {
                        worksheet.Cells[row, 1].Value = sv.MaSV;
                        worksheet.Cells[row, 2].Value = sv.HoTen;
                        worksheet.Cells[row, 3].Value = sv.NamSinh;
                        worksheet.Cells[row, 4].Value = sv.DiemTT;
                        worksheet.Cells[row, 5].Value = sv.DiemTB;

                        row++;
                    }
                }
                package.Save();
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}

class AVLtree
{   
    class Node
    {
        public SinhVien sv;
        public Node Left;
        public Node Right;

        public Node(SinhVien sv)
        {
            this.sv = sv;
        }
    }
    Node root;

    private Node RecursiveInsert(Node current, Node newNode)
    {
        if (current == null)
        {
            current = newNode;
            return current;
        }
        else if (string.Compare(newNode.sv.MaSV, current.sv.MaSV) < 0)
        {
            current.Left = RecursiveInsert(current.Left, newNode);
            current = BalanceTree(current);
        }
        else if (string.Compare(newNode.sv.MaSV, current.sv.MaSV) > 0)
        {
            current.Right = RecursiveInsert(current.Right, newNode);
            current = BalanceTree(current);
        }
        return current;
    }


    private Node RotateRight(Node y)
    {
        Node x = y.Left;
        Node T2 = x.Right;

        x.Right = y;
        y.Left = T2;
        return x;
    }

    private Node RotateLeft(Node x)
    {
        Node y = x.Right;
        Node T2 = y.Left;

        y.Left = x;
        x.Right = T2;

        return y;
    }

    private Node RotateRR(Node parent)
    {
        return RotateLeft(parent);
    }

    private Node RotateLL(Node parent)
    {
        return RotateRight(parent);
    }

    private Node RotateLR(Node parent)
    {
        parent.Left = RotateLeft(parent.Left);
        return RotateRight(parent);
    }

    private Node RotateRL(Node parent)
    {
        parent.Right = RotateRight(parent.Right);
        return RotateLeft(parent);
    }

    private int Max(int a, int b)
    {
        return (a > b) ? a : b;
    }

    private int GetHeight(Node current)
    {
        if (current == null)
        {
            return 0;
        }

        int leftHeight = GetHeight(current.Left);
        int rightHeight = GetHeight(current.Right);

        return Max(leftHeight, rightHeight) + 1;
    }

    private int BalanceFactor(Node current)  //Hệ số cân bằng
    {
        if (current == null)
        {
            return 0;
        }

        int leftHeight = GetHeight(current.Left);
        int rightHeight = GetHeight(current.Right);

        return leftHeight - rightHeight;
    }
    private Node BalanceTree(Node current)
    {
        int balance = BalanceFactor(current);

        if (balance > 1)
        {
            if (BalanceFactor(current.Left) >= 0)
            {
                current = RotateLL(current);
            }
            else
            {
                current = RotateLR(current);
            }
        }
        else if (balance < -1)
        {
            if (BalanceFactor(current.Right) <= 0)
            {
                current = RotateRR(current);
            }
            else
            {
                current = RotateRL(current);
            }
        }

        return current;
    }

    public void Add(SinhVien sv)       //Phương thức add sinh viên trong danh sách
    {
        if (Find(root, sv.MaSV) != null)
        {
            return;
        }
        Node newNode = new Node(sv);
        if (root == null)
        {
            root = newNode;
        }
        else
        {
            root = RecursiveInsert(root, newNode);
        }
    }

    public void Add1(SinhVien sv)     //Phương thức add cho người dùng (đã kiểm tra trùng bằng Find() từ trước
    {
        Node newNode = new Node(sv);
        if (root == null)
        {
            root = newNode;
        }
        else
        {
            root = RecursiveInsert(root, newNode);
            Console.WriteLine("Thêm sinh viên mới thành công");
        }
    }

    public SinhVien Find(string maSV)           
    {
        Node resultNode = Find(root, maSV);
        if (resultNode != null)
        {
            return resultNode.sv;
        }
        return null;
    }

    private Node Find(Node current, string maSV)
    {
        if (current == null)    //Không tìm thấy
        {
            return null;
        }

        if (string.Compare(maSV, current.sv.MaSV) == 0)        //Tìm thấy node
        {
            return current;
        }
        else if (string.Compare(maSV, current.sv.MaSV) < 0)  //Tìm tiếp cây con trái
        {
            return Find(current.Left, maSV);
        }
        else                                //Tìm tiếp cây con phải
        {
            return Find(current.Right, maSV);
        }
    }


    public void UpdateMark(string maSV, double mark)    //Tương tự find()
    {
        Node updated = Find(root, maSV);
        updated.sv.DiemTB = mark;
        Console.WriteLine("Cập nhật điểm cho sinh viên thành công");
    }

    public void Delete(string maSV)
    {
        root = Delete(root, maSV);
    }

    private Node Delete(Node current, string maSV)
    {
        if (current == null)
        {
            Console.WriteLine("Không tồn tại sinh viên với mã đã cho");
            return null;
        }

        if (string.Compare(maSV, current.sv.MaSV) < 0)   //Không phải node cần xóa thì tìm tiếp cây con bên trái
        {
            current.Left = Delete(current.Left, maSV);
            if (BalanceFactor(current) == -2)    //Cây lệch trái
            {
                if (BalanceFactor(current.Right) <= 0)
                {
                    current = RotateRight(current);
                }
                else
                {
                    current = RotateRL(current);
                }
            }
        }
        else if (string.Compare(maSV, current.sv.MaSV) > 0)   //Không phải node cần xóa thì tìm tiếp cây con bên phải
        {
            current.Right = Delete(current.Right, maSV);
            if (BalanceFactor(current) == 2)     //Cây lệch phải
            {
                if (BalanceFactor(current.Left) >= 0)
                {
                    current = RotateLeft(current);
                }
                else
                {
                    current = RotateLR(current);
                }
            }
        }
        else                        //Nút hiện tại là nút cần xóa
        {
            if (current.Right != null)   //Nếu nút cần xóa có cây con bên phải, tìm thế mạng
            {
                Node parent = current.Right;
                while (parent.Left != null)
                {
                    parent = parent.Left;
                }
                current.sv = parent.sv;
                current.Right = Delete(current.Right, parent.sv.MaSV);
                if (BalanceFactor(current) == 2)   //Xử lý trường hợp cây mất cân bằng sau xóa
                {
                    if (BalanceFactor(current.Left) >= 0)
                    {
                        current = RotateLeft(current);
                    }
                    else
                    {
                        current = RotateLR(current);
                    }
                }
            }
            else                //Nếu nút cần xóa không có cây con bên phải xóa nút và trả về cây con bên trái để thay thế.
            {
                Console.WriteLine("Xóa thành công! ");
                return current.Left;
            }
        }
        return current;
    }

    public void DisplayTree()     //Hiển thị cây duyệt theo thứ tự giữa
    {
        if (root == null)
        {
            Console.WriteLine("N");
            return;
        }
        InOrderTraversal(root);
    }

    private void InOrderTraversal(Node current)   //Duyệt thứ tự giữa
    {
        if (current != null)
        {
            InOrderTraversal(current.Left);
            Console.WriteLine("Mã sinh viên: " + current.sv.MaSV);
            Console.WriteLine("Họ tên: " + current.sv.HoTen);
            Console.WriteLine("Năm sinh: " + current.sv.NamSinh);
            Console.WriteLine("Điểm trúng tuyển: " + current.sv.DiemTT);
            Console.WriteLine("Điểm trung bình: " + current.sv.DiemTB);
            Console.WriteLine();
            InOrderTraversal(current.Right);
        }
    }

    public bool ExportFile(string sheetName)
    {
        using (ExcelPackage package = new ExcelPackage(new FileInfo(Path.Combine(Directory.GetCurrentDirectory(), "QLSinhVien.xlsx"))))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null)
            {
                worksheet = package.Workbook.Worksheets.Add(sheetName);
                // Tiêu đề cột
                worksheet.Cells[1, 1].Value = "MSSV";
                worksheet.Cells[1, 2].Value = "Họ tên";
                worksheet.Cells[1, 3].Value = "Năm sinh";
                worksheet.Cells[1, 4].Value = "Điểm trúng tuyển";
                worksheet.Cells[1, 5].Value = "Điểm trung bình";

                int row = 2; // Bắt đầu từ hàng 2 để viết dữ liệu
                ExportTreeToExcel(root, worksheet, ref row);
                package.Save();
                return true;
            }
            else
            {
                return false;
            }

        }
    }

    private void ExportTreeToExcel(Node current, ExcelWorksheet worksheet, ref int row)
    {
        if (current != null)
        {
            ExportTreeToExcel(current.Left, worksheet, ref row);

            // Ghi dữ liệu của node hiện tại vào Excel
            worksheet.Cells[row, 1].Value = current.sv.MaSV;
            worksheet.Cells[row, 2].Value = current.sv.HoTen;
            worksheet.Cells[row, 3].Value = current.sv.NamSinh;
            worksheet.Cells[row, 4].Value = current.sv.DiemTT;
            worksheet.Cells[row, 5].Value = current.sv.DiemTB;

            row++;

            ExportTreeToExcel(current.Right, worksheet, ref row);
        }
    }
}

class Program
{
    static void Main()
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        HashTable DSSV1 = new HashTable();
        AVLtree DSSV2 = new AVLtree();
        try
        {
            //Lấy dữ liệu từ sheet đầu tiên trong file
            var package = new ExcelPackage(new FileInfo(Path.Combine(Directory.GetCurrentDirectory(), "QLSinhVien.xlsx")));
            ExcelWorksheet workSheet = package.Workbook.Worksheets["Sheet1"];

            //Duyệt dữ liệu trừ dòng tiêu đề
            for (int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
            {
                string maSV = string.Empty, hoTen = string.Empty;
                int namSinh = 0;
                double diemTT = 0.0, diemTB = 0.0;
                int j = 1;

                var temp = workSheet.Cells[i, j++].Value;
                if (temp != null)
                {
                    maSV = temp.ToString();
                }

                temp = workSheet.Cells[i, j++].Value;
                if (temp != null)
                {
                    hoTen = temp.ToString();
                }

                temp = workSheet.Cells[i, j++].Value;
                if (temp != null)
                {
                    namSinh = Convert.ToInt32(temp.ToString());
                }

                temp = workSheet.Cells[i, j++].Value;
                if (temp != null)
                {
                    diemTT = Convert.ToDouble(temp.ToString());
                }

                temp = workSheet.Cells[i, j++].Value;
                if (temp != null)
                {
                    diemTB = Convert.ToDouble(temp.ToString());
                }

                SinhVien sv = new SinhVien(maSV, hoTen, namSinh, diemTT, diemTB);
                DSSV1.Add(sv);
                DSSV2.Add(sv);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
        }

        Console.WriteLine("***** Chương trình quản lý danh sách sinh viên *****\n----------------------------------------------");
        int n = 0;
        while(true)
        {
            Console.Write("Bạn muốn chọn bảng băm hay cây AVL cho việc thao tác (1 cho bảng băm, 2 cho cây AVL):  ");
            n = Convert.ToInt32(Console.ReadLine());
            if(n == 1 || n == 2)
            {
                break;
            }
            else
            {
                Console.WriteLine("Dữ liệu nhập không hợp lệ, vui lòng nhập lại!");
            }
        }
        
        if(n == 1)
        {
            Console.WriteLine("Bạn đã chọn CTDL bảng băm!\n-------------------------------------");
            while (true)
            {
                Console.WriteLine("Vui lòng chọn thao tác muốn thực hiện:");
                Console.WriteLine("   1. Thêm sinh viên vào danh sách");
                Console.WriteLine("   2. Tra cứu thông tin sinh viên");
                Console.WriteLine("   3. Cập nhật điểm của 1 sinh viên với mã xác định");
                Console.WriteLine("   4. Xóa 1 sinh viên với mã xác định");
                Console.WriteLine("   5. Hiển thị danh sách sinh viên hiện tại");
                Console.WriteLine("   6. Xuất danh sách sinh viên hiện tại ra file excel");
                Console.Write("Nhập lựa chọn mà bạn muốn thực hiện: ");
                int t = Convert.ToInt32(Console.ReadLine());
                switch(t)
                {
                    case 1:
                        {
                            string msv;
                            Console.WriteLine("Nhập thông tin sinh viên cần thêm:");
                            while (true)
                            {
                                Console.Write("Nhập mã sinh viên: ");
                                msv = Console.ReadLine();
                                if(DSSV1.Find(msv) != null)
                                {
                                    Console.WriteLine($"Sinh viên với mã {msv} đã tồn tại! Vui lòng nhập lại!");
                                }
                                else
                                {
                                    break;
                                }
                            }
                            Console.Write("Nhập họ và tên sinh viên cần thêm: ");
                            string hten = Console.ReadLine();
                            Console.Write("Nhập năm sinh: ");
                            int nsinh = Convert.ToInt32(Console.ReadLine());
                            Console.Write("Nhập điểm trúng tuyển: ");
                            double dtt = Convert.ToDouble(Console.ReadLine());
                            Console.Write("Nhập điểm trung bình: ");
                            double dtb = Convert.ToDouble(Console.ReadLine());
                            SinhVien nsv = new SinhVien(msv, hten, nsinh, dtt, dtb);
                            DSSV1.Add1(nsv);
                            Console.WriteLine("---------------------------------\n");
                            break;
                        }
                    case 2:
                        {
                            Console.Write("Nhập mã sinh viên cần tra cứu: ");
                            string tracuu = Console.ReadLine();
                            SinhVien sv = DSSV1.Find(tracuu);
                            if(sv != null)
                            {
                                Console.WriteLine($"Thông tin sinh viên với mã {tracuu} là: ");
                                Console.WriteLine("Họ tên: " + sv.HoTen);
                                Console.WriteLine("Năm sinh: " + sv.NamSinh);
                                Console.WriteLine("Điểm trúng tuyển: " + sv.DiemTT);
                                Console.WriteLine("Điểm trung bình: " + sv.DiemTB);
                            }
                            else
                            {
                                Console.WriteLine("Không tồn tại sinh viên với mã " + tracuu);
                            }
                            Console.WriteLine("---------------------------------\n");
                            break;
                        }
                    case 3:
                        {
                            string updateMaSV;
                            while (true)
                            {
                                Console.Write("Nhập mã sinh viên cần cập nhật điểm: ");
                                updateMaSV = Console.ReadLine();
                                if (DSSV1.Find(updateMaSV) != null)
                                {
                                    break;
                                }
                                else
                                {
                                    Console.WriteLine("Không tồn tại sinh viên với mã " + updateMaSV);
                                }
                            }
                            Console.Write("Nhập điểm trung bình cần cập nhật của sinh viên đã nhập: ");
                            double mark = Convert.ToDouble(Console.ReadLine());
                            DSSV1.UpdateMark(updateMaSV, mark);
                            Console.WriteLine("---------------------------------\n");
                            break;
                        }
                    case 4:
                        {
                            Console.Write("Nhập mã sinh viên của sinh viên cần xóa: ");
                            string deleteMaSV = Console.ReadLine();
                            DSSV1.Delete(deleteMaSV);
                            Console.WriteLine("---------------------------------\n");
                            break;
                        }
                    case 5:
                        {
                            Console.WriteLine("Danh sách sinh viên hiện tại là: ");
                            DSSV1.ShowAllStudent();
                            break;
                        }
                    case 6:
                        {
                            Console.Write("Nhập tên của sheet khi xuất DSSV: ");
                            string sheetName = Console.ReadLine();
                            if (DSSV1.ExportFile(sheetName))
                            {
                                Console.WriteLine("Xuất ra file excel thành công!");
                            }
                            else
                            {
                                Console.WriteLine("Tên sheet đã bị trùng! Vui lòng thực hiện lại");
                            }
                            break;
                        }
                    default:
                        {
                            Console.WriteLine("Dữ liệu nhập không hợp lệ!");
                            break;
                        }
                }
                Console.Write("Bạn có muốn tiếp tục thao tác không (1: yes, 2: no): ");
                int k = Convert.ToInt32(Console.ReadLine());
                if(k == 1)
                {
                    Console.WriteLine("Bạn đã chọn tiếp tục chương trình!");
                    Console.WriteLine("---------------------------------\n");
                }
                else
                {
                    Console.WriteLine("Bạn đã chọn thoát chương trình!\n--------------------------------------------");
                    break;
                }
            }
        }
        else
        {
            Console.WriteLine("Bạn đã chọn CTDL cây AVL!\n-------------------------------------");
            while (true)
            {
                Console.WriteLine("Vui lòng chọn thao tác muốn thực hiện:");
                Console.WriteLine("   1. Thêm sinh viên vào danh sách");
                Console.WriteLine("   2. Tra cứu thông tin sinh viên");
                Console.WriteLine("   3. Cập nhật điểm của 1 sinh viên với mã xác định");
                Console.WriteLine("   4. Xóa 1 sinh viên với mã xác định");
                Console.WriteLine("   5. Hiển thị danh sách sinh viên hiện tại");
                Console.WriteLine("   6. Xuất danh sách sinh viên hiện tại ra file excel");
                Console.Write("Nhập lựa chọn mà bạn muốn thực hiện: ");
                int t = Convert.ToInt32(Console.ReadLine());
                switch (t)
                {
                    case 1:
                        {
                            string msv;
                            Console.WriteLine("Nhập thông tin sinh viên cần thêm:");
                            while (true)
                            {
                                Console.Write("Nhập mã sinh viên: ");
                                msv = Console.ReadLine();
                                if (DSSV2.Find(msv) != null)
                                {
                                    Console.WriteLine($"Sinh viên với mã {msv} đã tồn tại! Vui lòng nhập lại!");
                                }
                                else
                                {
                                    break;
                                }
                            }
                            Console.Write("Nhập họ và tên sinh viên cần thêm: ");
                            string hten = Console.ReadLine();
                            Console.Write("Nhập năm sinh: ");
                            int nsinh = Convert.ToInt32(Console.ReadLine());
                            Console.Write("Nhập điểm trúng tuyển: ");
                            double dtt = Convert.ToDouble(Console.ReadLine());
                            Console.Write("Nhập điểm trung bình: ");
                            double dtb = Convert.ToDouble(Console.ReadLine());
                            SinhVien nsv = new SinhVien(msv, hten, nsinh, dtt, dtb);
                            DSSV1.Add1(nsv);
                            Console.WriteLine("---------------------------------\n");
                            break;
                        }
                    case 2:
                        {
                            Console.Write("Nhập mã sinh viên cần tra cứu: ");
                            string tracuu = Console.ReadLine();
                            SinhVien sv = DSSV2.Find(tracuu);
                            if (sv != null)
                            {
                                Console.WriteLine($"Thông tin sinh viên với mã {tracuu} là: ");
                                Console.WriteLine("Họ tên: " + sv.HoTen);
                                Console.WriteLine("Năm sinh: " + sv.NamSinh);
                                Console.WriteLine("Điểm trúng tuyển: " + sv.DiemTT);
                                Console.WriteLine("Điểm trung bình: " + sv.DiemTB);
                            }
                            else
                            {
                                Console.WriteLine("Không tồn tại sinh viên với mã " + tracuu);
                            }
                            Console.WriteLine("---------------------------------\n");
                            break;
                        }
                    case 3:
                        {
                            string updateMaSV;
                            while (true)
                            {
                                Console.Write("Nhập mã sinh viên cần cập nhật điểm: ");
                                updateMaSV = Console.ReadLine();
                                if (DSSV2.Find(updateMaSV) != null)
                                {
                                    break;
                                }
                                else
                                {
                                    Console.WriteLine("Không tồn tại sinh viên với mã " + updateMaSV);
                                }
                            }
                            Console.Write("Nhập điểm trung bình cần cập nhật của sinh viên đã nhập: ");
                            double mark = Convert.ToDouble(Console.ReadLine());
                            DSSV2.UpdateMark(updateMaSV, mark);
                            Console.WriteLine("---------------------------------\n");
                            break;
                        }
                    case 4:
                        {
                            Console.Write("Nhập mã sinh viên của sinh viên cần xóa: ");
                            string deleteMaSV = Console.ReadLine();
                            DSSV2.Delete(deleteMaSV);
                            Console.WriteLine("---------------------------------\n");
                            break;
                        }
                    case 5:
                        {
                            Console.WriteLine("Danh sách sinh viên hiện tại là: ");
                            DSSV2.DisplayTree();
                            break;
                        }
                    case 6:
                        {
                            Console.Write("Nhập tên của sheet khi xuất DSSV: ");
                            string sheetName = Console.ReadLine();
                            if (DSSV2.ExportFile(sheetName))
                            {
                                Console.WriteLine("Xuất ra file excel thành công!");
                            }
                            else
                            {
                                Console.WriteLine("Tên sheet đã bị trùng! Vui lòng thực hiện lại");
                            }
                            break;
                        }
                    default:
                        {
                            Console.WriteLine("Dữ liệu nhập không hợp lệ!");
                            break;
                        }
                }
                Console.Write("Bạn có muốn tiếp tục thao tác không (1: yes, 2: no): ");
                int k = Convert.ToInt32(Console.ReadLine());
                if (k == 1)
                {
                    Console.WriteLine("Bạn đã chọn tiếp tục chương trình!");
                    Console.WriteLine("---------------------------------\n");
                }
                else
                {
                    Console.WriteLine("Bạn đã chọn thoát chương trình!\n--------------------------------------------");
                    break;
                }
            }
        }

    }
}
