using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using file_transfer;
using Excel = Microsoft.Office.Interop.Excel;
using file_transfer.DTO;

public partial class Main : Form
{
    //This will hold our listener. We will only need to create one instance of this.
    private Listener listener;
    //This will hold our transfer client.
    private TransferClient transferClient;
    //This will hold our output folder.
    private string outputFolder;
    //This will hold our overall progress timer.
    private Timer tmrOverallProg;
    //This is our variable to determine of the server is running or not to accept another connection if our client
    //Disconnects
    private bool serverRunning;
    //Database Process
    private ExcelManagerCF db { get; set; }

    public Main()
    {
        InitializeComponent();

        //Init DB
        db = new ExcelManagerCF();

        //Create the listener and register the event.
        listener = new Listener();
        listener.Accepted += listener_Accepted;

        //Create the timer and register the event.
        tmrOverallProg = new Timer();
        tmrOverallProg.Interval = 1000;
        tmrOverallProg.Tick += tmrOverallProg_Tick;

        //Set our default output folder.
        outputFolder = "Transfers";

        //If it does not exist, create it.
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        btnConnect.Click += new EventHandler(btnConnect_Click);
        btnStartServer.Click += new EventHandler(btnStartServer_Click);
        btnStopServer.Click += new EventHandler(btnStopServer_Click);
        btnSendFile.Click += new EventHandler(btnSendFile_Click);
        btnPauseTransfer.Click += new EventHandler(btnPauseTransfer_Click);
        btnStopTransfer.Click += new EventHandler(btnStopTransfer_Click);
        btnOpenDir.Click += new EventHandler(btnOpenDir_Click);
        btnClearComplete.Click += new EventHandler(btnClearComplete_Click);
        btnExcelImport.Click += new EventHandler(btnExcelImport_Click);
        btnExcelExport.Click += new EventHandler(btnExcelExport_Click);
        btnStopServer.Enabled = false;
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        //Deregister all the events from the client if it is connected.
        deregisterEvents();
        base.OnFormClosing(e);
    }

    void tmrOverallProg_Tick(object sender, EventArgs e)
    {
        if (transferClient == null)
            return;
        //Get and display the overall progress.
        progressOverall.Value = transferClient.GetOverallProgress();
    }

    void listener_Accepted(object sender, SocketAcceptedEventArgs e)
    {
        if (InvokeRequired)
        {
            Invoke(new SocketAcceptedHandler(listener_Accepted), sender, e);
            return;
        }

        //Stop the listener
        listener.Stop();

        //Create the transfer client based on our newly connected socket.
        transferClient = new TransferClient(e.Accepted);
        //Set our output folder.
        transferClient.OutputFolder = outputFolder;
        //Register the events.
        registerEvents();
        //Run the client
        transferClient.Run();
        //Start the progress timer
        tmrOverallProg.Start();
        //And set the new connection state.
        setConnectionStatus(transferClient.EndPoint.Address.ToString());
    }

    private void btnConnect_Click(object sender, EventArgs e)
    {
        if (transferClient == null)
        {
            //Create our new transfer client.
            //And attempt to connect
            transferClient = new TransferClient();
            transferClient.Connect(txtCntHost.Text.Trim(), int.Parse(txtCntPort.Text.Trim()), connectCallback);
            Enabled = false;
        }
        else
        {
            //This means we're trying to disconnect.
            transferClient.Close();
            transferClient = null;
        }
    }

    private void connectCallback(object sender, string error)
    {
        if (InvokeRequired)
        {
            Invoke(new ConnectCallback(connectCallback), sender, error);
            return;
        }
        //Set the form to enabled.
        Enabled = true;
        //If the error is not equal to null, something went wrong.
        if (error != null)
        {
            transferClient.Close();
            transferClient = null;
            MessageBox.Show(error, "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        //Register the events
        registerEvents();
        //Set the output folder
        transferClient.OutputFolder = outputFolder;
        //Run the client
        transferClient.Run();
        //Set the connection status
        setConnectionStatus(transferClient.EndPoint.Address.ToString());
        //Start the progress timer.
        tmrOverallProg.Start();
        //Set our connect button text to "Disconnect"
        btnConnect.Text = "Disconnect";
    }

    private void registerEvents()
    {
        transferClient.Complete += transferClient_Complete;
        transferClient.Disconnected += transferClient_Disconnected;
        transferClient.ProgressChanged += transferClient_ProgressChanged;
        transferClient.Queued += transferClient_Queued;
        transferClient.Stopped += transferClient_Stopped;
    }

    void transferClient_Stopped(object sender, TransferQueue queue)
    {
        if (InvokeRequired)
        {
            Invoke(new TransferEventHandler(transferClient_Stopped), sender, queue);
            return;
        }
        //Remove the stopped transfer from view.
        lstTransfers.Items[queue.ID.ToString()].Remove();
    }

    void transferClient_Queued(object sender, TransferQueue queue)
    {
        if (InvokeRequired)
        {
            Invoke(new TransferEventHandler(transferClient_Queued), sender, queue);
            return;
        }

        //Create the LVI for the new transfer.
        ListViewItem i = new ListViewItem();
        i.Text = queue.ID.ToString();
        i.SubItems.Add(queue.Filename);
        //If the type equals download, it will use the string of "Download", if not, it'll use "Upload"
        i.SubItems.Add(queue.Type == QueueType.Download ? "Download" : "Upload");
        i.SubItems.Add("0%");
        i.Tag = queue; //Set the tag to queue so we can grab is easily.
        i.Name = queue.ID.ToString(); //Set the name of the item to the ID of our transfer for easy access.
        lstTransfers.Items.Add(i); //Add the item
        i.EnsureVisible();
        
        //If the type is download, let the uploader know we're ready.
        if (queue.Type == QueueType.Download)
        {
            transferClient.StartTransfer(queue);
        }
    }

    void transferClient_ProgressChanged(object sender, TransferQueue queue)
    {
        if (InvokeRequired)
        {
            Invoke(new TransferEventHandler(transferClient_ProgressChanged), sender, queue);
            return;
        }

        //Set the progress cell to our current progress.
        lstTransfers.Items[queue.ID.ToString()].SubItems[3].Text = queue.Progress + "%";
    }

    void transferClient_Disconnected(object sender, EventArgs e)
    {
        if (InvokeRequired)
        {
            Invoke(new EventHandler(transferClient_Disconnected), sender, e);
            return;
        }

        //Deregister the transfer client events
        deregisterEvents();

        //Close every transfer
        foreach (ListViewItem item in lstTransfers.Items)
        {
            TransferQueue queue = (TransferQueue)item.Tag;
            queue.Close();
        }
        //Clear the listview
        lstTransfers.Items.Clear();
        progressOverall.Value = 0;

        //Set the client to null
        transferClient = null;

        //Set the connection status to nothing
        setConnectionStatus("-");

        //If the server is still running, wait for another connection
        if (serverRunning)
        {
            listener.Start(int.Parse(txtServerPort.Text.Trim()));
            setConnectionStatus("Waiting...");
        }
        else //If we connected then disconnected, set the text back to connect.
        {
            btnConnect.Text = "Connect";
        }
    }

    void transferClient_Complete(object sender, TransferQueue queue)
    {
        //This just plays a little sound to let us know a transfer completed.
        System.Media.SystemSounds.Asterisk.Play();
    }

    private void deregisterEvents()
    {
        if (transferClient == null)
            return;
        transferClient.Complete -= transferClient_Complete;
        transferClient.Disconnected -= transferClient_Disconnected;
        transferClient.ProgressChanged -= transferClient_ProgressChanged;
        transferClient.Queued -= transferClient_Queued;
        transferClient.Stopped -= transferClient_Stopped;
    }

    private void setConnectionStatus(string connectedTo)
    {
        lblConnected.Text = "Connection: " + connectedTo;
    }

    private void btnStartServer_Click(object sender, EventArgs e)
    {
        //We disabled the button, but lets just do a quick check
        if (serverRunning)
            return;
        serverRunning = true;
        try
        {
            //Try to listen on the desired port
            listener.Start(int.Parse(txtServerPort.Text.Trim()));
            //Set the connection status to waiting
            setConnectionStatus("Waiting...");
            //Enable/Disable the server buttons.
            btnStartServer.Enabled = false;
            btnStopServer.Enabled = true;
        }
        catch
        {
            MessageBox.Show("Unable to listen on port " + txtServerPort.Text, "", MessageBoxButtons.OK, MessageBoxIcon.Error);

        }
    }

    private void btnStopServer_Click(object sender, EventArgs e)
    {
        if (!serverRunning)
            return;
        //Close the client if its active.
        if (transferClient != null)
        {
            transferClient.Close();
            //INSERT
            transferClient = null;
            //
        }
        //Stop the listener
        listener.Stop();
        //Stop the timer
        tmrOverallProg.Stop();
        //Reset the connection statis
        setConnectionStatus("-");
        //Set our variables and enable/disable the buttons.
        serverRunning = false;
        btnStartServer.Enabled = true;
        btnStopServer.Enabled = false;
    }

    private void btnClearComplete_Click(object sender, EventArgs e)
    {
        //Loop and clear all complete or inactive transfers
        foreach (ListViewItem i in lstTransfers.Items)
        {
            TransferQueue queue = (TransferQueue)i.Tag;

            if (queue.Progress == 100 || !queue.Running)
            {
                i.Remove();
            }
        }
    }

    private void btnOpenDir_Click(object sender, EventArgs e)
    {
        //Get a user defined save directory
        using (FolderBrowserDialog fb = new FolderBrowserDialog())
        {
            if (fb.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                outputFolder = fb.SelectedPath;

                if (transferClient != null)
                {
                    transferClient.OutputFolder = outputFolder;
                }

                txtSaveDir.Text = outputFolder;
            }
        }
    }

    private void btnSendFile_Click(object sender, EventArgs e)
    {
        if (transferClient == null)
            return;
        //Get the user desired files to send
        using (OpenFileDialog o = new OpenFileDialog())
        {
            o.Filter = "All Files (*.*)|*.*";
            o.Multiselect = true;

            if (o.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string file in o.FileNames)
                {
                    transferClient.QueueTransfer(file);
                }
            }
        }
    }

    private void btnPauseTransfer_Click(object sender, EventArgs e)
    {
        if (transferClient == null)
            return;
        //Loop and pause/resume all selected downloads.
        foreach (ListViewItem i in lstTransfers.SelectedItems)
        {
            TransferQueue queue = (TransferQueue)i.Tag;
            queue.Client.PauseTransfer(queue);
        }
    }

    private void btnStopTransfer_Click(object sender, EventArgs e)
    {
        if (transferClient == null)
            return;

        //Loop and stop all selected downloads.
        foreach (ListViewItem i in lstTransfers.SelectedItems)
        {
            TransferQueue queue = (TransferQueue)i.Tag;
            queue.Client.StopTransfer(queue);
            i.Remove();
        }

        progressOverall.Value = 0;
    }

    DataTable dt;
    List<Teacher> listTc;
    private void btnExcelExport_Click(Object sender, EventArgs e)
    {
        if (txtFile.Text != "")
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open(txtFile.Text);
            Excel._Worksheet sheet1 = workbook.Sheets[1];// Giám thị
            Excel._Worksheet sheet2 = workbook.Sheets[2];// Phòng thi
            Excel.Range x1Range = sheet1.UsedRange;
            Excel.Range x2Range = sheet2.UsedRange;

            int rowCountSheet1 = x1Range.Rows.Count;
            int columnCountSheet1 = x1Range.Columns.Count;

            int rowCountSheet2 = x2Range.Rows.Count;
            int columnCountSheet2 = x2Range.Columns.Count;

            List<Teacher> teachers = new List<Teacher>(rowCountSheet1);

            for (int i = 2; i <= rowCountSheet1; i++)
            {
                Teacher tc = new Teacher
                {
                    id = Int32.Parse(x1Range.Cells[i, 2].Value2.ToString()),
                    name = x1Range.Cells[i, 3].Value2.ToString(),
                    birthday = x1Range.Cells[i, 4].Value2.ToString(),
                    university = x1Range.Cells[i, 5].Value2.ToString()
                };
                teachers.Add(tc);
            }

            Random random = new Random();
            listTc = teachers.OrderBy(x => random.Next()).ToList();//==> list Teacher cần xử lý
            List<Room> rooms = new List<Room>(rowCountSheet2);//==> list phòng
            for (int i = 2; i <= rowCountSheet2; i++)
            {
                Room r = new Room
                {
                    id = Int32.Parse(x2Range.Cells[i, 1].Value2.ToString()),
                    nameRoom = x2Range.Cells[i, 2].Value2.ToString(),
                    localtion = x2Range.Cells[i, 3].Value2.ToString(),
                    note = x2Range.Cells[i, 4].Value2.ToString()
                };
                rooms.Add(r);
            }
            dt = new DataTable();
            dt.Columns.Add("ID phòng");
            dt.Columns.Add("Tên phòng");
            dt.Columns.Add("Địa điểm");
            dt.Columns.Add("ID CB1");
            dt.Columns.Add("Tên CB1");
            dt.Columns.Add("ID CB2");
            dt.Columns.Add("Tên CB2");//7
            for (int i = 0; i < rooms.Count; i++)
            {
                DataRow r = dt.NewRow();
                r[0] = rooms[i].id;
                r[1] = rooms[i].nameRoom;
                r[2] = rooms[i].localtion;
                r[3] = listTc[i * 2].id;
                r[4] = listTc[i * 2].name;
                r[5] = listTc[i * 2 + 1].id;
                r[6] = listTc[i * 2 + 1].name;

                dt.Rows.Add(r);
            }
            
            Export(dt, "PHÂN CÔNG CÁN BỘ COI THI", dt.Rows.Count/10);
        }
    }
    private void btnExcelImport_Click(object sender, EventArgs e)
    {
        using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*xlsx|Excel 97-2003 Workbook|*.xls" })
        {
            if(openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtFile.Text = openFileDialog.FileName;
            }
        }
    }
    private void Export(DataTable dt, string title, int sheetCount)
    {
        //Tao cac doi tuong Excel
        Excel.Application oExcel = new Excel.Application();
        Excel.Workbooks oBooks;
        Excel.Sheets oSheets;
        Excel.Workbook oBook;
        Excel.Worksheet[] oSheet = new Excel.Worksheet[sheetCount+1];
        //Tạo mới một Excel WorkBook 
        oExcel.Visible = true;
        oExcel.DisplayAlerts = false;
        oExcel.Application.SheetsInNewWorkbook = sheetCount+1;
        oBooks = oExcel.Workbooks;
        oBook = (Microsoft.Office.Interop.Excel.Workbook)(oExcel.Workbooks.Add(Type.Missing));
        oSheets = oBook.Worksheets;

        for(int i = 0; i < sheetCount; i++)
        {
            oSheet[i] = (Microsoft.Office.Interop.Excel.Worksheet)oSheets.get_Item(i+1);
            oSheet[i].Name = "sheet"+(i+1);
            DataTable datatable = new DataTable();
            datatable.Columns.Add("ID phòng");
            datatable.Columns.Add("Tên phòng");
            datatable.Columns.Add("Địa điểm");
            datatable.Columns.Add("ID CB1");
            datatable.Columns.Add("Tên CB1");
            datatable.Columns.Add("ID CB2");
            datatable.Columns.Add("Tên CB2");//7
            if (i == sheetCount - 1)
            {
                int lastIndex = dt.Rows.Count % 10;
                for (int j = 0; j < 10 - lastIndex; j++)
                {
                    DataRow row = datatable.NewRow();
                    row = dt.Rows[i * 10 + j];
                    datatable.ImportRow(row);
                }
            }
            else
            {
                for (int j = 0; j < 10; j++)
                {
                    DataRow row = datatable.NewRow();
                    row = dt.Rows[i * 10 + j];
                    datatable.ImportRow(row);
                }
            }
            fillContent1(oSheet[i], datatable, title);
        }


        oSheet[sheetCount] = (Microsoft.Office.Interop.Excel.Worksheet)oSheets.get_Item(sheetCount + 1);
        oSheet[sheetCount].Name = "sheet" + (sheetCount + 1);
        DataTable teachergs = new DataTable();
        teachergs.Columns.Add("Mã số");
        teachergs.Columns.Add("Tên cán bộ");
        teachergs.Columns.Add("Ngày Sinh");
        teachergs.Columns.Add("Cơ Quan");//4
        for (int i = sheetCount*20 + 1; i < listTc.Count; i++)
        {
            DataRow row = teachergs.NewRow();
            row[0] = listTc[i].id;
            row[1] = listTc[i].name;
            row[2] = listTc[i].birthday;
            row[3] = listTc[i].university;

            teachergs.Rows.Add(row);
        }
        fillContent2(oSheet[sheetCount], teachergs, "DANH SÁCH CÁN BỘ GIÁM SÁT");


    }

    private void fillContent1(Excel.Worksheet oSheet,DataTable dataTable , string title)
    {
        // Tạo phần đầu nếu muốn
        Microsoft.Office.Interop.Excel.Range tenTruong = oSheet.get_Range("A1", "C1");
        tenTruong.MergeCells = true;
        tenTruong.Value2 = "ĐẠI HỌC ĐÀ NẴNG";
        tenTruong.Font.Name = "Tahoma";
        tenTruong.Font.Size = "10";
        tenTruong.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

        Microsoft.Office.Interop.Excel.Range tenBK = oSheet.get_Range("A2", "C2");
        tenBK.MergeCells = true;
        tenBK.Value2 = "TRƯỜNG ĐẠI HỌC BÁCH KHOA";
        tenBK.Font.Name = "Tahoma";
        tenBK.Font.Underline = true;
        tenBK.Font.Size = "10";
        tenBK.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

        Microsoft.Office.Interop.Excel.Range congHoa = oSheet.get_Range("D1", "G1");
        congHoa.MergeCells = true;
        congHoa.Value2 = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
        congHoa.Font.Name = "Tahoma";
        congHoa.Font.Size = "10";
        congHoa.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

        Microsoft.Office.Interop.Excel.Range docLap = oSheet.get_Range("D2", "G2");
        docLap.MergeCells = true;
        docLap.Value2 = "Độc lập - Tự do - Hạnh phúc";
        docLap.Font.Name = "Tahoma";
        docLap.Font.Underline = true;
        docLap.Font.Size = "10";
        docLap.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

        Microsoft.Office.Interop.Excel.Range head = oSheet.get_Range("A4", "G4");
        head.MergeCells = true;
        head.Value2 = title;
        head.Font.Bold = true;
        head.Font.Name = "Tahoma";
        head.Font.Size = "18";
        head.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

        //Lop
        Microsoft.Office.Interop.Excel.Range lop = oSheet.get_Range("A7", "C7");
        lop.MergeCells = true;
        lop.Value2 = "LỚP : .............";
        lop.Font.Bold = false;
        lop.Font.Name = "Tahoma";
        lop.Font.Size = "10";
        lop.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
        //Giang Vien
        Microsoft.Office.Interop.Excel.Range giangVien = oSheet.get_Range("D7", "G7");
        giangVien.MergeCells = true;
        giangVien.Value2 = "GIẢNG VIÊN: .............";
        giangVien.Font.Name = "Tahoma";
        giangVien.Font.Size = "10";
        giangVien.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

        //Hoc Phan
        Microsoft.Office.Interop.Excel.Range hocPhan = oSheet.get_Range("A8", "C8");
        hocPhan.MergeCells = true;
        hocPhan.Value2 = "HỌC PHẦN: .............";
        hocPhan.Font.Name = "Tahoma";
        hocPhan.Font.Size = "10";
        hocPhan.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

        //ngay thi
        Microsoft.Office.Interop.Excel.Range ngayThi = oSheet.get_Range("D8", "G8");
        ngayThi.MergeCells = true;
        ngayThi.Value2 = "NGÀY THI: .............";
        ngayThi.Font.Name = "Tahoma";
        ngayThi.Font.Size = "10";
        ngayThi.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

        //Phong Dao Tao
        Microsoft.Office.Interop.Excel.Range phongDaoTao = oSheet.get_Range("A9", "C9");
        phongDaoTao.MergeCells = true;
        phongDaoTao.Value2 = "PHÒNG ĐÀO TẠO: .............";
        phongDaoTao.Font.Name = "Tahoma";
        phongDaoTao.Font.Size = "10";
        phongDaoTao.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

        //Phong Thi
        Microsoft.Office.Interop.Excel.Range phongThi = oSheet.get_Range("D9", "G9");
        phongThi.MergeCells = true;
        phongThi.Value2 = "PHÒNG THI: .............";
        phongThi.Font.Name = "Tahoma";
        phongThi.Font.Size = "10";
        phongThi.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

        // Tạo tiêu đề cột 
        Microsoft.Office.Interop.Excel.Range cl1 = oSheet.get_Range("A11", "A11");
        cl1.Value2 = "Mã Phòng";

        Microsoft.Office.Interop.Excel.Range cl2 = oSheet.get_Range("B11", "B11");
        cl2.Value2 = "Tên Phòng";
        cl2.ColumnWidth = 13.5;

        Microsoft.Office.Interop.Excel.Range cl3 = oSheet.get_Range("C11", "C11");
        cl3.Value2 = "Địa điểm";
        cl3.ColumnWidth = 13.5;

        Microsoft.Office.Interop.Excel.Range cl4 = oSheet.get_Range("D11", "D11");
        cl4.Value2 = "Mã số CB1";
        cl4.ColumnWidth = 13.5;

        Microsoft.Office.Interop.Excel.Range cl5 = oSheet.get_Range("E11", "E11");
        cl5.Value2 = "Tên CB1";
        cl5.ColumnWidth = 20.0;

        Microsoft.Office.Interop.Excel.Range cl6 = oSheet.get_Range("F11", "F11");
        cl6.Value2 = "Mã số CB2";
        cl6.ColumnWidth = 13.5;

        Microsoft.Office.Interop.Excel.Range cl7 = oSheet.get_Range("G11", "G11");
        cl7.Value2 = "Tên CB2";
        cl7.ColumnWidth = 20.0;

        Microsoft.Office.Interop.Excel.Range rowHead = oSheet.get_Range("A11", "G11");
        rowHead.Font.Bold = true;
        // Kẻ viền
        rowHead.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlSolid;

        // Thiết lập màu nền
        rowHead.Interior.ColorIndex = 15;
        rowHead.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

        // Tạo mẳng đối tượng để lưu dữ toàn bồ dữ liệu trong DataTable,
        // vì dữ liệu được được gán vào các Cell trong Excel phải thông qua object thuần.

        object[,] arr = new object[dataTable.Rows.Count, dataTable.Columns.Count];

        //Chuyển dữ liệu từ DataTable vào mảng đối tượng
        for (int r = 0; r < dataTable.Rows.Count; r++)
        {
            DataRow dr = dataTable.Rows[r];
            for (int c = 0; c < dataTable.Columns.Count; c++)
            {
                arr[r, c] = dr[c];
            }
        }
        //Thiết lập vùng điền dữ liệu
        int rowStart = 12;
        int columnStart = 1;
        int rowEnd = rowStart + dataTable.Rows.Count - 1;
        int columnEnd = dataTable.Columns.Count;

        // Ô bắt đầu điền dữ liệu
        Microsoft.Office.Interop.Excel.Range c1 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[rowStart, columnStart];

        // Ô kết thúc điền dữ liệu
        Microsoft.Office.Interop.Excel.Range c2 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[rowEnd, columnEnd];

        // Lấy về vùng điền dữ liệu
        Microsoft.Office.Interop.Excel.Range range = oSheet.get_Range(c1, c2);

        //Điền dữ liệu vào vùng đã thiết lập
        range.Value2 = arr;

        // Kẻ viền
        range.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlSolid;

        // Căn giữa cột STT
        //Microsoft.Office.Interop.Excel.Range c3 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[rowEnd, columnStart];
        //Microsoft.Office.Interop.Excel.Range c4 = oSheet.get_Range(c1, c3);
        //oSheet.get_Range(c3, c4).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        //Microsoft.Office.Interop.Excel.Range truongKhoa = oSheet.get_Range("A40", "B40");
        //truongKhoa.MergeCells = true;
        //truongKhoa.Value2 = "TRƯỞNG KHOA/BỘ MÔN";
        //truongKhoa.Font.Name = "Tahoma";
        //truongKhoa.Font.Size = "10";
        //truongKhoa.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

        //Microsoft.Office.Interop.Excel.Range daNang = oSheet.get_Range("D40", "G40");
        //daNang.MergeCells = true;
        //daNang.Value2 = "ĐÀ NẴNG, Ngày.......Tháng.......Năm.....";
        //daNang.Font.Name = "Tahoma";
        //daNang.Font.Size = "10";
        //daNang.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

        //Microsoft.Office.Interop.Excel.Range canBo = oSheet.get_Range("C41", "E41");
        //canBo.MergeCells = true;
        //canBo.Value2 = "CÁN BỘ COI THI";
        //canBo.Font.Name = "Tahoma";
        //canBo.Font.Size = "10";
        //canBo.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
    }
    private void fillContent2(Excel.Worksheet oSheet, DataTable dataTable, string title)
    {
        // Tạo phần đầu nếu muốn
        Microsoft.Office.Interop.Excel.Range tenTruong = oSheet.get_Range("A1", "B1");
        tenTruong.MergeCells = true;
        tenTruong.Value2 = "ĐẠI HỌC ĐÀ NẴNG";
        tenTruong.Font.Name = "Tahoma";
        tenTruong.Font.Size = "10";
        tenTruong.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

        Microsoft.Office.Interop.Excel.Range tenBK = oSheet.get_Range("A2", "B2");
        tenBK.MergeCells = true;
        tenBK.Value2 = "TRƯỜNG ĐẠI HỌC BÁCH KHOA";
        tenBK.Font.Name = "Tahoma";
        tenBK.Font.Underline = true;
        tenBK.Font.Size = "10";
        tenBK.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

        Microsoft.Office.Interop.Excel.Range congHoa = oSheet.get_Range("C1", "D1");
        congHoa.MergeCells = true;
        congHoa.Value2 = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
        congHoa.Font.Name = "Tahoma";
        congHoa.Font.Size = "10";
        congHoa.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

        Microsoft.Office.Interop.Excel.Range docLap = oSheet.get_Range("C2", "D2");
        docLap.MergeCells = true;
        docLap.Value2 = "Độc lập - Tự do - Hạnh phúc";
        docLap.Font.Name = "Tahoma";
        docLap.Font.Underline = true;
        docLap.Font.Size = "10";
        docLap.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

        Microsoft.Office.Interop.Excel.Range head = oSheet.get_Range("A4", "D4");
        head.MergeCells = true;
        head.Value2 = title;
        head.Font.Bold = true;
        head.Font.Name = "Tahoma";
        head.Font.Size = "18";
        head.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

        //Lop
        Microsoft.Office.Interop.Excel.Range lop = oSheet.get_Range("A7", "B7");
        lop.MergeCells = true;
        lop.Value2 = "LỚP : .............";
        lop.Font.Bold = false;
        lop.Font.Name = "Tahoma";
        lop.Font.Size = "10";
        lop.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
        //Giang Vien
        Microsoft.Office.Interop.Excel.Range giangVien = oSheet.get_Range("C7", "D7");
        giangVien.MergeCells = true;
        giangVien.Value2 = "GIẢNG VIÊN: .............";
        giangVien.Font.Name = "Tahoma";
        giangVien.Font.Size = "10";
        giangVien.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

        //Hoc Phan
        Microsoft.Office.Interop.Excel.Range hocPhan = oSheet.get_Range("A8", "B8");
        hocPhan.MergeCells = true;
        hocPhan.Value2 = "HỌC PHẦN: .............";
        hocPhan.Font.Name = "Tahoma";
        hocPhan.Font.Size = "10";
        hocPhan.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

        //ngay thi
        Microsoft.Office.Interop.Excel.Range ngayThi = oSheet.get_Range("C8", "D8");
        ngayThi.MergeCells = true;
        ngayThi.Value2 = "NGÀY THI: .............";
        ngayThi.Font.Name = "Tahoma";
        ngayThi.Font.Size = "10";
        ngayThi.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

        //Phong Dao Tao
        Microsoft.Office.Interop.Excel.Range phongDaoTao = oSheet.get_Range("A9", "B9");
        phongDaoTao.MergeCells = true;
        phongDaoTao.Value2 = "PHÒNG ĐÀO TẠO: .............";
        phongDaoTao.Font.Name = "Tahoma";
        phongDaoTao.Font.Size = "10";
        phongDaoTao.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

        //Phong Thi
        Microsoft.Office.Interop.Excel.Range phongThi = oSheet.get_Range("C9", "D9");
        phongThi.MergeCells = true;
        phongThi.Value2 = "PHÒNG THI: .............";
        phongThi.Font.Name = "Tahoma";
        phongThi.Font.Size = "10";
        phongThi.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

        // Tạo tiêu đề cột 
        Microsoft.Office.Interop.Excel.Range cl1 = oSheet.get_Range("A11", "A11");
        cl1.Value2 = "Mã Cán Bộ";
        cl1.ColumnWidth = 20;

        Microsoft.Office.Interop.Excel.Range cl2 = oSheet.get_Range("B11", "B11");
        cl2.Value2 = "Tên Cán Bộ";
        cl2.ColumnWidth = 20;

        Microsoft.Office.Interop.Excel.Range cl3 = oSheet.get_Range("C11", "C11");
        cl3.Value2 = "Ngày Sinh";
        cl3.ColumnWidth = 20;

        Microsoft.Office.Interop.Excel.Range cl4 = oSheet.get_Range("D11", "D11");
        cl4.Value2 = "Cơ Quan";
        cl4.ColumnWidth = 20;

        Microsoft.Office.Interop.Excel.Range rowHead = oSheet.get_Range("A11", "D11");
        rowHead.Font.Bold = true;
        // Kẻ viền
        rowHead.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlSolid;

        // Thiết lập màu nền
        rowHead.Interior.ColorIndex = 15;
        rowHead.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

        // Tạo mẳng đối tượng để lưu dữ toàn bồ dữ liệu trong DataTable,
        // vì dữ liệu được được gán vào các Cell trong Excel phải thông qua object thuần.

        object[,] arr = new object[dataTable.Rows.Count, dataTable.Columns.Count];

        //Chuyển dữ liệu từ DataTable vào mảng đối tượng
        for (int r = 0; r < dataTable.Rows.Count; r++)
        {
            DataRow dr = dataTable.Rows[r];
            for (int c = 0; c < dataTable.Columns.Count; c++)
            {
                arr[r, c] = dr[c];
            }
        }
        //Thiết lập vùng điền dữ liệu
        int rowStart = 12;
        int columnStart = 1;
        int rowEnd = rowStart + dataTable.Rows.Count - 1;
        int columnEnd = dataTable.Columns.Count;

        // Ô bắt đầu điền dữ liệu
        Microsoft.Office.Interop.Excel.Range c1 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[rowStart, columnStart];

        // Ô kết thúc điền dữ liệu
        Microsoft.Office.Interop.Excel.Range c2 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[rowEnd, columnEnd];

        // Lấy về vùng điền dữ liệu
        Microsoft.Office.Interop.Excel.Range range = oSheet.get_Range(c1, c2);

        //Điền dữ liệu vào vùng đã thiết lập
        range.Value2 = arr;

        // Kẻ viền
        range.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlSolid;

        // Căn giữa cột STT
        //Microsoft.Office.Interop.Excel.Range c3 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[rowEnd, columnStart];
        //Microsoft.Office.Interop.Excel.Range c4 = oSheet.get_Range(c1, c3);
        //oSheet.get_Range(c3, c4).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        //Microsoft.Office.Interop.Excel.Range truongKhoa = oSheet.get_Range("A40", "B40");
        //truongKhoa.MergeCells = true;
        //truongKhoa.Value2 = "TRƯỞNG KHOA/BỘ MÔN";
        //truongKhoa.Font.Name = "Tahoma";
        //truongKhoa.Font.Size = "10";
        //truongKhoa.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

        //Microsoft.Office.Interop.Excel.Range daNang = oSheet.get_Range("D40", "G40");
        //daNang.MergeCells = true;
        //daNang.Value2 = "ĐÀ NẴNG, Ngày.......Tháng.......Năm.....";
        //daNang.Font.Name = "Tahoma";
        //daNang.Font.Size = "10";
        //daNang.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

        //Microsoft.Office.Interop.Excel.Range canBo = oSheet.get_Range("C41", "E41");
        //canBo.MergeCells = true;
        //canBo.Value2 = "CÁN BỘ COI THI";
        //canBo.Font.Name = "Tahoma";
        //canBo.Font.Size = "10";
        //canBo.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
    }
}