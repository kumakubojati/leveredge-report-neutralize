<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Main
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main))
        Me.TabCon = New System.Windows.Forms.TabControl()
        Me.TabOutMas = New System.Windows.Forms.TabPage()
        Me.PicBar = New System.Windows.Forms.PictureBox()
        Me.prgbar = New System.Windows.Forms.Label()
        Me.btnBrow_OutMasDes = New System.Windows.Forms.Button()
        Me.txtDest_OutMas = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnNeuOutMas = New System.Windows.Forms.Button()
        Me.btnbrow_outmassrc = New System.Windows.Forms.Button()
        Me.txtoutmas_src = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TPProdMast = New System.Windows.Forms.TabPage()
        Me.PicBar_ProdMas = New System.Windows.Forms.PictureBox()
        Me.btnBrowProdMas_dest = New System.Windows.Forms.Button()
        Me.txtProdMas_dest = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.btnNeu_ProdMas = New System.Windows.Forms.Button()
        Me.btnBrowProdMas_src = New System.Windows.Forms.Button()
        Me.txtProdMas_src = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.OFD_src = New System.Windows.Forms.OpenFileDialog()
        Me.SFD_Dest = New System.Windows.Forms.SaveFileDialog()
        Me.BWOutMas = New System.ComponentModel.BackgroundWorker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TabConSales = New System.Windows.Forms.TabControl()
        Me.LP3 = New System.Windows.Forms.TabPage()
        Me.GBLP3 = New System.Windows.Forms.GroupBox()
        Me.RBQty = New System.Windows.Forms.RadioButton()
        Me.RBRupiah = New System.Windows.Forms.RadioButton()
        Me.RBAll = New System.Windows.Forms.RadioButton()
        Me.PicBarLP3 = New System.Windows.Forms.PictureBox()
        Me.btnBrow_LP3_Dest = New System.Windows.Forms.Button()
        Me.txtLP3_dest = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btnNeuLP3 = New System.Windows.Forms.Button()
        Me.btnBrow_LP3_src = New System.Windows.Forms.Button()
        Me.txtLP3_src = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TPSalesPerfRep = New System.Windows.Forms.TabPage()
        Me.gbRepType_SPR = New System.Windows.Forms.GroupBox()
        Me.RBSPR_Avg = New System.Windows.Forms.RadioButton()
        Me.RBSPR_Val = New System.Windows.Forms.RadioButton()
        Me.RBSPR_Qty = New System.Windows.Forms.RadioButton()
        Me.PicBar_SPR = New System.Windows.Forms.PictureBox()
        Me.btnBrow_SPR_dest = New System.Windows.Forms.Button()
        Me.txtSPR_dest = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnNeu_SPR = New System.Windows.Forms.Button()
        Me.btnBrow_SPR_src = New System.Windows.Forms.Button()
        Me.txtSPR_src = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.DSStab = New System.Windows.Forms.TabPage()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.rbDSSprod = New System.Windows.Forms.RadioButton()
        Me.rbDSSrupiah = New System.Windows.Forms.RadioButton()
        Me.rbDSSall = New System.Windows.Forms.RadioButton()
        Me.PicBar_DSS = New System.Windows.Forms.PictureBox()
        Me.btnBrowDSS_dest = New System.Windows.Forms.Button()
        Me.txtDSS_dest = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.btnNeu_DSS = New System.Windows.Forms.Button()
        Me.bntBrowDSS_src = New System.Windows.Forms.Button()
        Me.txtDSS_src = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.BWLP3 = New System.ComponentModel.BackgroundWorker()
        Me.TabConStock = New System.Windows.Forms.TabControl()
        Me.TPDistStock = New System.Windows.Forms.TabPage()
        Me.PicBar_DistStock = New System.Windows.Forms.PictureBox()
        Me.btnBrowDistStock_Dest = New System.Windows.Forms.Button()
        Me.txtDistStock_Dest = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.btnNeu_DistStock = New System.Windows.Forms.Button()
        Me.btnBrowDistStock_src = New System.Windows.Forms.Button()
        Me.txtDistStock_src = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.tpDailyStockMutation = New System.Windows.Forms.TabPage()
        Me.PicBar_DSM = New System.Windows.Forms.PictureBox()
        Me.btnBrow_DSM_dest = New System.Windows.Forms.Button()
        Me.txtDSM_dest = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.btnNeu_DSM = New System.Windows.Forms.Button()
        Me.btnBrow_DSM_src = New System.Windows.Forms.Button()
        Me.txtDSM_src = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.BWDistStock = New System.ComponentModel.BackgroundWorker()
        Me.BWProdMas = New System.ComponentModel.BackgroundWorker()
        Me.BWSalesPerf = New System.ComponentModel.BackgroundWorker()
        Me.BWDSM = New System.ComponentModel.BackgroundWorker()
        Me.tbListPromo = New System.Windows.Forms.TabPage()
        Me.PicBar_Promo = New System.Windows.Forms.PictureBox()
        Me.btnBrowPromo_dest = New System.Windows.Forms.Button()
        Me.txtProm_dest = New System.Windows.Forms.TextBox()
        Me.txtPromo_src = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.btnNeu_Promo = New System.Windows.Forms.Button()
        Me.btnBrowProm_src = New System.Windows.Forms.Button()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.tcListPromo = New System.Windows.Forms.TabControl()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.bwDSS = New System.ComponentModel.BackgroundWorker()
        Me.bwPromo = New System.ComponentModel.BackgroundWorker()
        Me.TabCon.SuspendLayout()
        Me.TabOutMas.SuspendLayout()
        CType(Me.PicBar, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TPProdMast.SuspendLayout()
        CType(Me.PicBar_ProdMas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabConSales.SuspendLayout()
        Me.LP3.SuspendLayout()
        Me.GBLP3.SuspendLayout()
        CType(Me.PicBarLP3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TPSalesPerfRep.SuspendLayout()
        Me.gbRepType_SPR.SuspendLayout()
        CType(Me.PicBar_SPR, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.DSStab.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.PicBar_DSS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabConStock.SuspendLayout()
        Me.TPDistStock.SuspendLayout()
        CType(Me.PicBar_DistStock, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tpDailyStockMutation.SuspendLayout()
        CType(Me.PicBar_DSM, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbListPromo.SuspendLayout()
        CType(Me.PicBar_Promo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tcListPromo.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabCon
        '
        Me.TabCon.Controls.Add(Me.TabOutMas)
        Me.TabCon.Controls.Add(Me.TPProdMast)
        Me.TabCon.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabCon.Location = New System.Drawing.Point(17, 27)
        Me.TabCon.Name = "TabCon"
        Me.TabCon.SelectedIndex = 0
        Me.TabCon.Size = New System.Drawing.Size(413, 183)
        Me.TabCon.TabIndex = 0
        '
        'TabOutMas
        '
        Me.TabOutMas.BackColor = System.Drawing.Color.White
        Me.TabOutMas.Controls.Add(Me.PicBar)
        Me.TabOutMas.Controls.Add(Me.prgbar)
        Me.TabOutMas.Controls.Add(Me.btnBrow_OutMasDes)
        Me.TabOutMas.Controls.Add(Me.txtDest_OutMas)
        Me.TabOutMas.Controls.Add(Me.Label2)
        Me.TabOutMas.Controls.Add(Me.btnNeuOutMas)
        Me.TabOutMas.Controls.Add(Me.btnbrow_outmassrc)
        Me.TabOutMas.Controls.Add(Me.txtoutmas_src)
        Me.TabOutMas.Controls.Add(Me.Label1)
        Me.TabOutMas.Location = New System.Drawing.Point(4, 23)
        Me.TabOutMas.Name = "TabOutMas"
        Me.TabOutMas.Padding = New System.Windows.Forms.Padding(3)
        Me.TabOutMas.Size = New System.Drawing.Size(405, 156)
        Me.TabOutMas.TabIndex = 0
        Me.TabOutMas.Text = "Outlet Master"
        '
        'PicBar
        '
        Me.PicBar.Image = CType(resources.GetObject("PicBar.Image"), System.Drawing.Image)
        Me.PicBar.Location = New System.Drawing.Point(5, 72)
        Me.PicBar.Name = "PicBar"
        Me.PicBar.Size = New System.Drawing.Size(80, 80)
        Me.PicBar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar.TabIndex = 9
        Me.PicBar.TabStop = False
        Me.PicBar.Visible = False
        '
        'prgbar
        '
        Me.prgbar.AutoSize = True
        Me.prgbar.Font = New System.Drawing.Font("Leelawadee UI Semilight", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.prgbar.Location = New System.Drawing.Point(91, 131)
        Me.prgbar.Name = "prgbar"
        Me.prgbar.Size = New System.Drawing.Size(49, 21)
        Me.prgbar.TabIndex = 8
        Me.prgbar.Text = "%bar"
        Me.prgbar.Visible = False
        '
        'btnBrow_OutMasDes
        '
        Me.btnBrow_OutMasDes.BackColor = System.Drawing.Color.Transparent
        Me.btnBrow_OutMasDes.Image = CType(resources.GetObject("btnBrow_OutMasDes.Image"), System.Drawing.Image)
        Me.btnBrow_OutMasDes.Location = New System.Drawing.Point(313, 45)
        Me.btnBrow_OutMasDes.Name = "btnBrow_OutMasDes"
        Me.btnBrow_OutMasDes.Size = New System.Drawing.Size(25, 23)
        Me.btnBrow_OutMasDes.TabIndex = 6
        Me.btnBrow_OutMasDes.UseVisualStyleBackColor = False
        '
        'txtDest_OutMas
        '
        Me.txtDest_OutMas.Location = New System.Drawing.Point(79, 46)
        Me.txtDest_OutMas.Name = "txtDest_OutMas"
        Me.txtDest_OutMas.Size = New System.Drawing.Size(229, 22)
        Me.txtDest_OutMas.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 51)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(70, 14)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Destination"
        '
        'btnNeuOutMas
        '
        Me.btnNeuOutMas.Enabled = False
        Me.btnNeuOutMas.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNeuOutMas.Image = CType(resources.GetObject("btnNeuOutMas.Image"), System.Drawing.Image)
        Me.btnNeuOutMas.Location = New System.Drawing.Point(344, 12)
        Me.btnNeuOutMas.Name = "btnNeuOutMas"
        Me.btnNeuOutMas.Size = New System.Drawing.Size(55, 56)
        Me.btnNeuOutMas.TabIndex = 3
        Me.btnNeuOutMas.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnNeuOutMas.UseVisualStyleBackColor = True
        '
        'btnbrow_outmassrc
        '
        Me.btnbrow_outmassrc.BackColor = System.Drawing.Color.Transparent
        Me.btnbrow_outmassrc.Image = CType(resources.GetObject("btnbrow_outmassrc.Image"), System.Drawing.Image)
        Me.btnbrow_outmassrc.Location = New System.Drawing.Point(313, 11)
        Me.btnbrow_outmassrc.Name = "btnbrow_outmassrc"
        Me.btnbrow_outmassrc.Size = New System.Drawing.Size(25, 23)
        Me.btnbrow_outmassrc.TabIndex = 2
        Me.btnbrow_outmassrc.UseVisualStyleBackColor = False
        '
        'txtoutmas_src
        '
        Me.txtoutmas_src.Location = New System.Drawing.Point(79, 12)
        Me.txtoutmas_src.Name = "txtoutmas_src"
        Me.txtoutmas_src.Size = New System.Drawing.Size(229, 22)
        Me.txtoutmas_src.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 14)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Source"
        '
        'TPProdMast
        '
        Me.TPProdMast.Controls.Add(Me.PicBar_ProdMas)
        Me.TPProdMast.Controls.Add(Me.btnBrowProdMas_dest)
        Me.TPProdMast.Controls.Add(Me.txtProdMas_dest)
        Me.TPProdMast.Controls.Add(Me.Label11)
        Me.TPProdMast.Controls.Add(Me.btnNeu_ProdMas)
        Me.TPProdMast.Controls.Add(Me.btnBrowProdMas_src)
        Me.TPProdMast.Controls.Add(Me.txtProdMas_src)
        Me.TPProdMast.Controls.Add(Me.Label12)
        Me.TPProdMast.Location = New System.Drawing.Point(4, 23)
        Me.TPProdMast.Name = "TPProdMast"
        Me.TPProdMast.Padding = New System.Windows.Forms.Padding(3)
        Me.TPProdMast.Size = New System.Drawing.Size(405, 156)
        Me.TPProdMast.TabIndex = 1
        Me.TPProdMast.Text = "Product Master"
        Me.TPProdMast.UseVisualStyleBackColor = True
        '
        'PicBar_ProdMas
        '
        Me.PicBar_ProdMas.Image = CType(resources.GetObject("PicBar_ProdMas.Image"), System.Drawing.Image)
        Me.PicBar_ProdMas.Location = New System.Drawing.Point(5, 69)
        Me.PicBar_ProdMas.Name = "PicBar_ProdMas"
        Me.PicBar_ProdMas.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_ProdMas.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_ProdMas.TabIndex = 18
        Me.PicBar_ProdMas.TabStop = False
        Me.PicBar_ProdMas.Visible = False
        '
        'btnBrowProdMas_dest
        '
        Me.btnBrowProdMas_dest.BackColor = System.Drawing.Color.Transparent
        Me.btnBrowProdMas_dest.Image = CType(resources.GetObject("btnBrowProdMas_dest.Image"), System.Drawing.Image)
        Me.btnBrowProdMas_dest.Location = New System.Drawing.Point(313, 42)
        Me.btnBrowProdMas_dest.Name = "btnBrowProdMas_dest"
        Me.btnBrowProdMas_dest.Size = New System.Drawing.Size(25, 23)
        Me.btnBrowProdMas_dest.TabIndex = 16
        Me.btnBrowProdMas_dest.UseVisualStyleBackColor = False
        '
        'txtProdMas_dest
        '
        Me.txtProdMas_dest.Location = New System.Drawing.Point(79, 43)
        Me.txtProdMas_dest.Name = "txtProdMas_dest"
        Me.txtProdMas_dest.Size = New System.Drawing.Size(229, 22)
        Me.txtProdMas_dest.TabIndex = 15
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(6, 48)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(70, 14)
        Me.Label11.TabIndex = 14
        Me.Label11.Text = "Destination"
        '
        'btnNeu_ProdMas
        '
        Me.btnNeu_ProdMas.Enabled = False
        Me.btnNeu_ProdMas.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNeu_ProdMas.Image = CType(resources.GetObject("btnNeu_ProdMas.Image"), System.Drawing.Image)
        Me.btnNeu_ProdMas.Location = New System.Drawing.Point(344, 9)
        Me.btnNeu_ProdMas.Name = "btnNeu_ProdMas"
        Me.btnNeu_ProdMas.Size = New System.Drawing.Size(55, 56)
        Me.btnNeu_ProdMas.TabIndex = 13
        Me.btnNeu_ProdMas.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnNeu_ProdMas.UseVisualStyleBackColor = True
        '
        'btnBrowProdMas_src
        '
        Me.btnBrowProdMas_src.BackColor = System.Drawing.Color.Transparent
        Me.btnBrowProdMas_src.Image = CType(resources.GetObject("btnBrowProdMas_src.Image"), System.Drawing.Image)
        Me.btnBrowProdMas_src.Location = New System.Drawing.Point(313, 8)
        Me.btnBrowProdMas_src.Name = "btnBrowProdMas_src"
        Me.btnBrowProdMas_src.Size = New System.Drawing.Size(25, 23)
        Me.btnBrowProdMas_src.TabIndex = 12
        Me.btnBrowProdMas_src.UseVisualStyleBackColor = False
        '
        'txtProdMas_src
        '
        Me.txtProdMas_src.Location = New System.Drawing.Point(79, 9)
        Me.txtProdMas_src.Name = "txtProdMas_src"
        Me.txtProdMas_src.Size = New System.Drawing.Size(229, 22)
        Me.txtProdMas_src.TabIndex = 11
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(6, 12)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(43, 14)
        Me.Label12.TabIndex = 10
        Me.Label12.Text = "Source"
        '
        'OFD_src
        '
        Me.OFD_src.FileName = "Source Excel"
        Me.OFD_src.Filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx"
        Me.OFD_src.RestoreDirectory = True
        Me.OFD_src.Title = "Select Your Excel File"
        '
        'SFD_Dest
        '
        Me.SFD_Dest.FileName = "Neutralize_OM_"
        Me.SFD_Dest.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*xlsx"
        Me.SFD_Dest.RestoreDirectory = True
        '
        'BWOutMas
        '
        Me.BWOutMas.WorkerReportsProgress = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Calibri", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(18, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(50, 17)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Master"
        '
        'TabConSales
        '
        Me.TabConSales.Controls.Add(Me.LP3)
        Me.TabConSales.Controls.Add(Me.TPSalesPerfRep)
        Me.TabConSales.Controls.Add(Me.DSStab)
        Me.TabConSales.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabConSales.Location = New System.Drawing.Point(19, 237)
        Me.TabConSales.Name = "TabConSales"
        Me.TabConSales.SelectedIndex = 0
        Me.TabConSales.Size = New System.Drawing.Size(411, 236)
        Me.TabConSales.TabIndex = 2
        '
        'LP3
        '
        Me.LP3.BackColor = System.Drawing.Color.White
        Me.LP3.Controls.Add(Me.GBLP3)
        Me.LP3.Controls.Add(Me.PicBarLP3)
        Me.LP3.Controls.Add(Me.btnBrow_LP3_Dest)
        Me.LP3.Controls.Add(Me.txtLP3_dest)
        Me.LP3.Controls.Add(Me.Label6)
        Me.LP3.Controls.Add(Me.btnNeuLP3)
        Me.LP3.Controls.Add(Me.btnBrow_LP3_src)
        Me.LP3.Controls.Add(Me.txtLP3_src)
        Me.LP3.Controls.Add(Me.Label5)
        Me.LP3.Location = New System.Drawing.Point(4, 23)
        Me.LP3.Name = "LP3"
        Me.LP3.Padding = New System.Windows.Forms.Padding(3)
        Me.LP3.Size = New System.Drawing.Size(403, 209)
        Me.LP3.TabIndex = 0
        Me.LP3.Text = "Weekly Stock And Sales (LP3)"
        '
        'GBLP3
        '
        Me.GBLP3.Controls.Add(Me.RBQty)
        Me.GBLP3.Controls.Add(Me.RBRupiah)
        Me.GBLP3.Controls.Add(Me.RBAll)
        Me.GBLP3.Location = New System.Drawing.Point(4, 6)
        Me.GBLP3.Name = "GBLP3"
        Me.GBLP3.Size = New System.Drawing.Size(251, 53)
        Me.GBLP3.TabIndex = 13
        Me.GBLP3.TabStop = False
        Me.GBLP3.Text = "Report Type"
        '
        'RBQty
        '
        Me.RBQty.AutoSize = True
        Me.RBQty.Location = New System.Drawing.Point(146, 22)
        Me.RBQty.Name = "RBQty"
        Me.RBQty.Size = New System.Drawing.Size(102, 18)
        Me.RBQty.TabIndex = 2
        Me.RBQty.TabStop = True
        Me.RBQty.Text = "Detail Per SKU"
        Me.RBQty.UseVisualStyleBackColor = True
        '
        'RBRupiah
        '
        Me.RBRupiah.AutoSize = True
        Me.RBRupiah.Location = New System.Drawing.Point(67, 22)
        Me.RBRupiah.Name = "RBRupiah"
        Me.RBRupiah.Size = New System.Drawing.Size(64, 18)
        Me.RBRupiah.TabIndex = 1
        Me.RBRupiah.TabStop = True
        Me.RBRupiah.Text = "Rupiah"
        Me.RBRupiah.UseVisualStyleBackColor = True
        '
        'RBAll
        '
        Me.RBAll.AutoSize = True
        Me.RBAll.Location = New System.Drawing.Point(7, 22)
        Me.RBAll.Name = "RBAll"
        Me.RBAll.Size = New System.Drawing.Size(40, 18)
        Me.RBAll.TabIndex = 0
        Me.RBAll.TabStop = True
        Me.RBAll.Text = "All"
        Me.RBAll.UseVisualStyleBackColor = True
        '
        'PicBarLP3
        '
        Me.PicBarLP3.Image = CType(resources.GetObject("PicBarLP3.Image"), System.Drawing.Image)
        Me.PicBarLP3.Location = New System.Drawing.Point(5, 125)
        Me.PicBarLP3.Name = "PicBarLP3"
        Me.PicBarLP3.Size = New System.Drawing.Size(80, 80)
        Me.PicBarLP3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBarLP3.TabIndex = 10
        Me.PicBarLP3.TabStop = False
        Me.PicBarLP3.Visible = False
        '
        'btnBrow_LP3_Dest
        '
        Me.btnBrow_LP3_Dest.Image = CType(resources.GetObject("btnBrow_LP3_Dest.Image"), System.Drawing.Image)
        Me.btnBrow_LP3_Dest.Location = New System.Drawing.Point(310, 97)
        Me.btnBrow_LP3_Dest.Name = "btnBrow_LP3_Dest"
        Me.btnBrow_LP3_Dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_LP3_Dest.TabIndex = 12
        Me.btnBrow_LP3_Dest.UseVisualStyleBackColor = True
        '
        'txtLP3_dest
        '
        Me.txtLP3_dest.Location = New System.Drawing.Point(76, 98)
        Me.txtLP3_dest.Name = "txtLP3_dest"
        Me.txtLP3_dest.Size = New System.Drawing.Size(229, 22)
        Me.txtLP3_dest.TabIndex = 11
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 103)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(70, 14)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Destination"
        '
        'btnNeuLP3
        '
        Me.btnNeuLP3.Enabled = False
        Me.btnNeuLP3.Image = CType(resources.GetObject("btnNeuLP3.Image"), System.Drawing.Image)
        Me.btnNeuLP3.Location = New System.Drawing.Point(343, 61)
        Me.btnNeuLP3.Name = "btnNeuLP3"
        Me.btnNeuLP3.Size = New System.Drawing.Size(54, 59)
        Me.btnNeuLP3.TabIndex = 9
        Me.btnNeuLP3.UseVisualStyleBackColor = True
        '
        'btnBrow_LP3_src
        '
        Me.btnBrow_LP3_src.Image = CType(resources.GetObject("btnBrow_LP3_src.Image"), System.Drawing.Image)
        Me.btnBrow_LP3_src.Location = New System.Drawing.Point(310, 63)
        Me.btnBrow_LP3_src.Name = "btnBrow_LP3_src"
        Me.btnBrow_LP3_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_LP3_src.TabIndex = 8
        Me.btnBrow_LP3_src.UseVisualStyleBackColor = True
        '
        'txtLP3_src
        '
        Me.txtLP3_src.Location = New System.Drawing.Point(76, 64)
        Me.txtLP3_src.Name = "txtLP3_src"
        Me.txtLP3_src.Size = New System.Drawing.Size(229, 22)
        Me.txtLP3_src.TabIndex = 7
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 69)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(43, 14)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Source"
        '
        'TPSalesPerfRep
        '
        Me.TPSalesPerfRep.Controls.Add(Me.gbRepType_SPR)
        Me.TPSalesPerfRep.Controls.Add(Me.PicBar_SPR)
        Me.TPSalesPerfRep.Controls.Add(Me.btnBrow_SPR_dest)
        Me.TPSalesPerfRep.Controls.Add(Me.txtSPR_dest)
        Me.TPSalesPerfRep.Controls.Add(Me.Label7)
        Me.TPSalesPerfRep.Controls.Add(Me.btnNeu_SPR)
        Me.TPSalesPerfRep.Controls.Add(Me.btnBrow_SPR_src)
        Me.TPSalesPerfRep.Controls.Add(Me.txtSPR_src)
        Me.TPSalesPerfRep.Controls.Add(Me.Label13)
        Me.TPSalesPerfRep.Location = New System.Drawing.Point(4, 23)
        Me.TPSalesPerfRep.Name = "TPSalesPerfRep"
        Me.TPSalesPerfRep.Padding = New System.Windows.Forms.Padding(3)
        Me.TPSalesPerfRep.Size = New System.Drawing.Size(403, 209)
        Me.TPSalesPerfRep.TabIndex = 1
        Me.TPSalesPerfRep.Text = "Sales Performance Report"
        Me.TPSalesPerfRep.UseVisualStyleBackColor = True
        '
        'gbRepType_SPR
        '
        Me.gbRepType_SPR.Controls.Add(Me.RBSPR_Avg)
        Me.gbRepType_SPR.Controls.Add(Me.RBSPR_Val)
        Me.gbRepType_SPR.Controls.Add(Me.RBSPR_Qty)
        Me.gbRepType_SPR.Location = New System.Drawing.Point(5, 5)
        Me.gbRepType_SPR.Name = "gbRepType_SPR"
        Me.gbRepType_SPR.Size = New System.Drawing.Size(228, 53)
        Me.gbRepType_SPR.TabIndex = 22
        Me.gbRepType_SPR.TabStop = False
        Me.gbRepType_SPR.Text = "Report Type"
        '
        'RBSPR_Avg
        '
        Me.RBSPR_Avg.AutoSize = True
        Me.RBSPR_Avg.Enabled = False
        Me.RBSPR_Avg.Location = New System.Drawing.Point(148, 22)
        Me.RBSPR_Avg.Name = "RBSPR_Avg"
        Me.RBSPR_Avg.Size = New System.Drawing.Size(68, 18)
        Me.RBSPR_Avg.TabIndex = 2
        Me.RBSPR_Avg.TabStop = True
        Me.RBSPR_Avg.Text = "Average"
        Me.RBSPR_Avg.UseVisualStyleBackColor = True
        '
        'RBSPR_Val
        '
        Me.RBSPR_Val.AutoSize = True
        Me.RBSPR_Val.Location = New System.Drawing.Point(85, 22)
        Me.RBSPR_Val.Name = "RBSPR_Val"
        Me.RBSPR_Val.Size = New System.Drawing.Size(56, 18)
        Me.RBSPR_Val.TabIndex = 1
        Me.RBSPR_Val.TabStop = True
        Me.RBSPR_Val.Text = "Value"
        Me.RBSPR_Val.UseVisualStyleBackColor = True
        '
        'RBSPR_Qty
        '
        Me.RBSPR_Qty.AutoSize = True
        Me.RBSPR_Qty.Enabled = False
        Me.RBSPR_Qty.Location = New System.Drawing.Point(7, 22)
        Me.RBSPR_Qty.Name = "RBSPR_Qty"
        Me.RBSPR_Qty.Size = New System.Drawing.Size(70, 18)
        Me.RBSPR_Qty.TabIndex = 0
        Me.RBSPR_Qty.TabStop = True
        Me.RBSPR_Qty.Text = "Quantity"
        Me.RBSPR_Qty.UseVisualStyleBackColor = True
        '
        'PicBar_SPR
        '
        Me.PicBar_SPR.Image = CType(resources.GetObject("PicBar_SPR.Image"), System.Drawing.Image)
        Me.PicBar_SPR.Location = New System.Drawing.Point(6, 124)
        Me.PicBar_SPR.Name = "PicBar_SPR"
        Me.PicBar_SPR.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_SPR.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_SPR.TabIndex = 18
        Me.PicBar_SPR.TabStop = False
        Me.PicBar_SPR.Visible = False
        '
        'btnBrow_SPR_dest
        '
        Me.btnBrow_SPR_dest.Image = CType(resources.GetObject("btnBrow_SPR_dest.Image"), System.Drawing.Image)
        Me.btnBrow_SPR_dest.Location = New System.Drawing.Point(311, 96)
        Me.btnBrow_SPR_dest.Name = "btnBrow_SPR_dest"
        Me.btnBrow_SPR_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_SPR_dest.TabIndex = 21
        Me.btnBrow_SPR_dest.UseVisualStyleBackColor = True
        '
        'txtSPR_dest
        '
        Me.txtSPR_dest.Location = New System.Drawing.Point(77, 97)
        Me.txtSPR_dest.Name = "txtSPR_dest"
        Me.txtSPR_dest.Size = New System.Drawing.Size(229, 22)
        Me.txtSPR_dest.TabIndex = 20
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(4, 102)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(70, 14)
        Me.Label7.TabIndex = 19
        Me.Label7.Text = "Destination"
        '
        'btnNeu_SPR
        '
        Me.btnNeu_SPR.Enabled = False
        Me.btnNeu_SPR.Image = CType(resources.GetObject("btnNeu_SPR.Image"), System.Drawing.Image)
        Me.btnNeu_SPR.Location = New System.Drawing.Point(344, 60)
        Me.btnNeu_SPR.Name = "btnNeu_SPR"
        Me.btnNeu_SPR.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_SPR.TabIndex = 17
        Me.btnNeu_SPR.UseVisualStyleBackColor = True
        '
        'btnBrow_SPR_src
        '
        Me.btnBrow_SPR_src.Image = CType(resources.GetObject("btnBrow_SPR_src.Image"), System.Drawing.Image)
        Me.btnBrow_SPR_src.Location = New System.Drawing.Point(311, 62)
        Me.btnBrow_SPR_src.Name = "btnBrow_SPR_src"
        Me.btnBrow_SPR_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrow_SPR_src.TabIndex = 16
        Me.btnBrow_SPR_src.UseVisualStyleBackColor = True
        '
        'txtSPR_src
        '
        Me.txtSPR_src.Location = New System.Drawing.Point(77, 63)
        Me.txtSPR_src.Name = "txtSPR_src"
        Me.txtSPR_src.Size = New System.Drawing.Size(229, 22)
        Me.txtSPR_src.TabIndex = 15
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(7, 68)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(43, 14)
        Me.Label13.TabIndex = 14
        Me.Label13.Text = "Source"
        '
        'DSStab
        '
        Me.DSStab.Controls.Add(Me.GroupBox2)
        Me.DSStab.Controls.Add(Me.PicBar_DSS)
        Me.DSStab.Controls.Add(Me.btnBrowDSS_dest)
        Me.DSStab.Controls.Add(Me.txtDSS_dest)
        Me.DSStab.Controls.Add(Me.Label19)
        Me.DSStab.Controls.Add(Me.btnNeu_DSS)
        Me.DSStab.Controls.Add(Me.bntBrowDSS_src)
        Me.DSStab.Controls.Add(Me.txtDSS_src)
        Me.DSStab.Controls.Add(Me.Label20)
        Me.DSStab.Location = New System.Drawing.Point(4, 23)
        Me.DSStab.Name = "DSStab"
        Me.DSStab.Padding = New System.Windows.Forms.Padding(3)
        Me.DSStab.Size = New System.Drawing.Size(403, 209)
        Me.DSStab.TabIndex = 2
        Me.DSStab.Text = "Daily Sales Summary"
        Me.DSStab.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rbDSSprod)
        Me.GroupBox2.Controls.Add(Me.rbDSSrupiah)
        Me.GroupBox2.Controls.Add(Me.rbDSSall)
        Me.GroupBox2.Location = New System.Drawing.Point(5, 5)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(251, 53)
        Me.GroupBox2.TabIndex = 22
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Report Type"
        '
        'rbDSSprod
        '
        Me.rbDSSprod.AutoSize = True
        Me.rbDSSprod.Location = New System.Drawing.Point(146, 22)
        Me.rbDSSprod.Name = "rbDSSprod"
        Me.rbDSSprod.Size = New System.Drawing.Size(62, 18)
        Me.rbDSSprod.TabIndex = 2
        Me.rbDSSprod.TabStop = True
        Me.rbDSSprod.Text = "Produk"
        Me.rbDSSprod.UseVisualStyleBackColor = True
        '
        'rbDSSrupiah
        '
        Me.rbDSSrupiah.AutoSize = True
        Me.rbDSSrupiah.Location = New System.Drawing.Point(67, 22)
        Me.rbDSSrupiah.Name = "rbDSSrupiah"
        Me.rbDSSrupiah.Size = New System.Drawing.Size(64, 18)
        Me.rbDSSrupiah.TabIndex = 1
        Me.rbDSSrupiah.TabStop = True
        Me.rbDSSrupiah.Text = "Rupiah"
        Me.rbDSSrupiah.UseVisualStyleBackColor = True
        '
        'rbDSSall
        '
        Me.rbDSSall.AutoSize = True
        Me.rbDSSall.Checked = True
        Me.rbDSSall.Location = New System.Drawing.Point(7, 22)
        Me.rbDSSall.Name = "rbDSSall"
        Me.rbDSSall.Size = New System.Drawing.Size(40, 18)
        Me.rbDSSall.TabIndex = 0
        Me.rbDSSall.TabStop = True
        Me.rbDSSall.Text = "All"
        Me.rbDSSall.UseVisualStyleBackColor = True
        '
        'PicBar_DSS
        '
        Me.PicBar_DSS.Image = CType(resources.GetObject("PicBar_DSS.Image"), System.Drawing.Image)
        Me.PicBar_DSS.Location = New System.Drawing.Point(6, 124)
        Me.PicBar_DSS.Name = "PicBar_DSS"
        Me.PicBar_DSS.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_DSS.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_DSS.TabIndex = 18
        Me.PicBar_DSS.TabStop = False
        Me.PicBar_DSS.Visible = False
        '
        'btnBrowDSS_dest
        '
        Me.btnBrowDSS_dest.Image = CType(resources.GetObject("btnBrowDSS_dest.Image"), System.Drawing.Image)
        Me.btnBrowDSS_dest.Location = New System.Drawing.Point(311, 96)
        Me.btnBrowDSS_dest.Name = "btnBrowDSS_dest"
        Me.btnBrowDSS_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrowDSS_dest.TabIndex = 21
        Me.btnBrowDSS_dest.UseVisualStyleBackColor = True
        '
        'txtDSS_dest
        '
        Me.txtDSS_dest.Location = New System.Drawing.Point(77, 97)
        Me.txtDSS_dest.Name = "txtDSS_dest"
        Me.txtDSS_dest.Size = New System.Drawing.Size(229, 22)
        Me.txtDSS_dest.TabIndex = 20
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(4, 102)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(70, 14)
        Me.Label19.TabIndex = 19
        Me.Label19.Text = "Destination"
        '
        'btnNeu_DSS
        '
        Me.btnNeu_DSS.Enabled = False
        Me.btnNeu_DSS.Image = CType(resources.GetObject("btnNeu_DSS.Image"), System.Drawing.Image)
        Me.btnNeu_DSS.Location = New System.Drawing.Point(344, 60)
        Me.btnNeu_DSS.Name = "btnNeu_DSS"
        Me.btnNeu_DSS.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_DSS.TabIndex = 17
        Me.btnNeu_DSS.UseVisualStyleBackColor = True
        '
        'bntBrowDSS_src
        '
        Me.bntBrowDSS_src.Image = CType(resources.GetObject("bntBrowDSS_src.Image"), System.Drawing.Image)
        Me.bntBrowDSS_src.Location = New System.Drawing.Point(311, 62)
        Me.bntBrowDSS_src.Name = "bntBrowDSS_src"
        Me.bntBrowDSS_src.Size = New System.Drawing.Size(26, 23)
        Me.bntBrowDSS_src.TabIndex = 16
        Me.bntBrowDSS_src.UseVisualStyleBackColor = True
        '
        'txtDSS_src
        '
        Me.txtDSS_src.Location = New System.Drawing.Point(77, 63)
        Me.txtDSS_src.Name = "txtDSS_src"
        Me.txtDSS_src.Size = New System.Drawing.Size(229, 22)
        Me.txtDSS_src.TabIndex = 15
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(7, 68)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(43, 14)
        Me.Label20.TabIndex = 14
        Me.Label20.Text = "Source"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Calibri", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(22, 219)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(38, 17)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Sales"
        '
        'BWLP3
        '
        Me.BWLP3.WorkerReportsProgress = True
        '
        'TabConStock
        '
        Me.TabConStock.Controls.Add(Me.TPDistStock)
        Me.TabConStock.Controls.Add(Me.tpDailyStockMutation)
        Me.TabConStock.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabConStock.Location = New System.Drawing.Point(438, 26)
        Me.TabConStock.Name = "TabConStock"
        Me.TabConStock.SelectedIndex = 0
        Me.TabConStock.Size = New System.Drawing.Size(413, 184)
        Me.TabConStock.TabIndex = 4
        '
        'TPDistStock
        '
        Me.TPDistStock.BackColor = System.Drawing.Color.White
        Me.TPDistStock.Controls.Add(Me.PicBar_DistStock)
        Me.TPDistStock.Controls.Add(Me.btnBrowDistStock_Dest)
        Me.TPDistStock.Controls.Add(Me.txtDistStock_Dest)
        Me.TPDistStock.Controls.Add(Me.Label8)
        Me.TPDistStock.Controls.Add(Me.btnNeu_DistStock)
        Me.TPDistStock.Controls.Add(Me.btnBrowDistStock_src)
        Me.TPDistStock.Controls.Add(Me.txtDistStock_src)
        Me.TPDistStock.Controls.Add(Me.Label9)
        Me.TPDistStock.Location = New System.Drawing.Point(4, 23)
        Me.TPDistStock.Name = "TPDistStock"
        Me.TPDistStock.Padding = New System.Windows.Forms.Padding(3)
        Me.TPDistStock.Size = New System.Drawing.Size(405, 157)
        Me.TPDistStock.TabIndex = 0
        Me.TPDistStock.Text = "Distributor Stock Report"
        '
        'PicBar_DistStock
        '
        Me.PicBar_DistStock.Image = CType(resources.GetObject("PicBar_DistStock.Image"), System.Drawing.Image)
        Me.PicBar_DistStock.Location = New System.Drawing.Point(6, 75)
        Me.PicBar_DistStock.Name = "PicBar_DistStock"
        Me.PicBar_DistStock.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_DistStock.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_DistStock.TabIndex = 9
        Me.PicBar_DistStock.TabStop = False
        Me.PicBar_DistStock.Visible = False
        '
        'btnBrowDistStock_Dest
        '
        Me.btnBrowDistStock_Dest.BackColor = System.Drawing.Color.Transparent
        Me.btnBrowDistStock_Dest.Image = CType(resources.GetObject("btnBrowDistStock_Dest.Image"), System.Drawing.Image)
        Me.btnBrowDistStock_Dest.Location = New System.Drawing.Point(313, 45)
        Me.btnBrowDistStock_Dest.Name = "btnBrowDistStock_Dest"
        Me.btnBrowDistStock_Dest.Size = New System.Drawing.Size(25, 23)
        Me.btnBrowDistStock_Dest.TabIndex = 6
        Me.btnBrowDistStock_Dest.UseVisualStyleBackColor = False
        '
        'txtDistStock_Dest
        '
        Me.txtDistStock_Dest.Location = New System.Drawing.Point(79, 46)
        Me.txtDistStock_Dest.Name = "txtDistStock_Dest"
        Me.txtDistStock_Dest.Size = New System.Drawing.Size(229, 22)
        Me.txtDistStock_Dest.TabIndex = 5
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 51)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(70, 14)
        Me.Label8.TabIndex = 4
        Me.Label8.Text = "Destination"
        '
        'btnNeu_DistStock
        '
        Me.btnNeu_DistStock.Enabled = False
        Me.btnNeu_DistStock.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNeu_DistStock.Image = CType(resources.GetObject("btnNeu_DistStock.Image"), System.Drawing.Image)
        Me.btnNeu_DistStock.Location = New System.Drawing.Point(344, 12)
        Me.btnNeu_DistStock.Name = "btnNeu_DistStock"
        Me.btnNeu_DistStock.Size = New System.Drawing.Size(55, 56)
        Me.btnNeu_DistStock.TabIndex = 3
        Me.btnNeu_DistStock.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnNeu_DistStock.UseVisualStyleBackColor = True
        '
        'btnBrowDistStock_src
        '
        Me.btnBrowDistStock_src.BackColor = System.Drawing.Color.Transparent
        Me.btnBrowDistStock_src.Image = CType(resources.GetObject("btnBrowDistStock_src.Image"), System.Drawing.Image)
        Me.btnBrowDistStock_src.Location = New System.Drawing.Point(313, 11)
        Me.btnBrowDistStock_src.Name = "btnBrowDistStock_src"
        Me.btnBrowDistStock_src.Size = New System.Drawing.Size(25, 23)
        Me.btnBrowDistStock_src.TabIndex = 2
        Me.btnBrowDistStock_src.UseVisualStyleBackColor = False
        '
        'txtDistStock_src
        '
        Me.txtDistStock_src.Location = New System.Drawing.Point(79, 12)
        Me.txtDistStock_src.Name = "txtDistStock_src"
        Me.txtDistStock_src.Size = New System.Drawing.Size(229, 22)
        Me.txtDistStock_src.TabIndex = 1
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(6, 15)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(43, 14)
        Me.Label9.TabIndex = 0
        Me.Label9.Text = "Source"
        '
        'tpDailyStockMutation
        '
        Me.tpDailyStockMutation.Controls.Add(Me.PicBar_DSM)
        Me.tpDailyStockMutation.Controls.Add(Me.btnBrow_DSM_dest)
        Me.tpDailyStockMutation.Controls.Add(Me.txtDSM_dest)
        Me.tpDailyStockMutation.Controls.Add(Me.Label14)
        Me.tpDailyStockMutation.Controls.Add(Me.btnNeu_DSM)
        Me.tpDailyStockMutation.Controls.Add(Me.btnBrow_DSM_src)
        Me.tpDailyStockMutation.Controls.Add(Me.txtDSM_src)
        Me.tpDailyStockMutation.Controls.Add(Me.Label15)
        Me.tpDailyStockMutation.Location = New System.Drawing.Point(4, 23)
        Me.tpDailyStockMutation.Name = "tpDailyStockMutation"
        Me.tpDailyStockMutation.Padding = New System.Windows.Forms.Padding(3)
        Me.tpDailyStockMutation.Size = New System.Drawing.Size(405, 157)
        Me.tpDailyStockMutation.TabIndex = 1
        Me.tpDailyStockMutation.Text = "Daily Stock Mutation"
        Me.tpDailyStockMutation.UseVisualStyleBackColor = True
        '
        'PicBar_DSM
        '
        Me.PicBar_DSM.Image = CType(resources.GetObject("PicBar_DSM.Image"), System.Drawing.Image)
        Me.PicBar_DSM.Location = New System.Drawing.Point(6, 75)
        Me.PicBar_DSM.Name = "PicBar_DSM"
        Me.PicBar_DSM.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_DSM.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_DSM.TabIndex = 17
        Me.PicBar_DSM.TabStop = False
        Me.PicBar_DSM.Visible = False
        '
        'btnBrow_DSM_dest
        '
        Me.btnBrow_DSM_dest.BackColor = System.Drawing.Color.Transparent
        Me.btnBrow_DSM_dest.Image = CType(resources.GetObject("btnBrow_DSM_dest.Image"), System.Drawing.Image)
        Me.btnBrow_DSM_dest.Location = New System.Drawing.Point(313, 45)
        Me.btnBrow_DSM_dest.Name = "btnBrow_DSM_dest"
        Me.btnBrow_DSM_dest.Size = New System.Drawing.Size(25, 23)
        Me.btnBrow_DSM_dest.TabIndex = 16
        Me.btnBrow_DSM_dest.UseVisualStyleBackColor = False
        '
        'txtDSM_dest
        '
        Me.txtDSM_dest.Location = New System.Drawing.Point(79, 46)
        Me.txtDSM_dest.Name = "txtDSM_dest"
        Me.txtDSM_dest.Size = New System.Drawing.Size(229, 22)
        Me.txtDSM_dest.TabIndex = 15
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(6, 51)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(70, 14)
        Me.Label14.TabIndex = 14
        Me.Label14.Text = "Destination"
        '
        'btnNeu_DSM
        '
        Me.btnNeu_DSM.Enabled = False
        Me.btnNeu_DSM.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNeu_DSM.Image = CType(resources.GetObject("btnNeu_DSM.Image"), System.Drawing.Image)
        Me.btnNeu_DSM.Location = New System.Drawing.Point(344, 12)
        Me.btnNeu_DSM.Name = "btnNeu_DSM"
        Me.btnNeu_DSM.Size = New System.Drawing.Size(55, 56)
        Me.btnNeu_DSM.TabIndex = 13
        Me.btnNeu_DSM.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnNeu_DSM.UseVisualStyleBackColor = True
        '
        'btnBrow_DSM_src
        '
        Me.btnBrow_DSM_src.BackColor = System.Drawing.Color.Transparent
        Me.btnBrow_DSM_src.Image = CType(resources.GetObject("btnBrow_DSM_src.Image"), System.Drawing.Image)
        Me.btnBrow_DSM_src.Location = New System.Drawing.Point(313, 11)
        Me.btnBrow_DSM_src.Name = "btnBrow_DSM_src"
        Me.btnBrow_DSM_src.Size = New System.Drawing.Size(25, 23)
        Me.btnBrow_DSM_src.TabIndex = 12
        Me.btnBrow_DSM_src.UseVisualStyleBackColor = False
        '
        'txtDSM_src
        '
        Me.txtDSM_src.Location = New System.Drawing.Point(79, 12)
        Me.txtDSM_src.Name = "txtDSM_src"
        Me.txtDSM_src.Size = New System.Drawing.Size(229, 22)
        Me.txtDSM_src.TabIndex = 11
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(6, 15)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(43, 14)
        Me.Label15.TabIndex = 10
        Me.Label15.Text = "Source"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.BackColor = System.Drawing.Color.Transparent
        Me.Label10.Font = New System.Drawing.Font("Calibri", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(436, 9)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(41, 17)
        Me.Label10.TabIndex = 5
        Me.Label10.Text = "Stock"
        '
        'BWDistStock
        '
        '
        'BWProdMas
        '
        '
        'BWSalesPerf
        '
        '
        'BWDSM
        '
        '
        'tbListPromo
        '
        Me.tbListPromo.BackColor = System.Drawing.Color.White
        Me.tbListPromo.Controls.Add(Me.PicBar_Promo)
        Me.tbListPromo.Controls.Add(Me.btnBrowPromo_dest)
        Me.tbListPromo.Controls.Add(Me.txtProm_dest)
        Me.tbListPromo.Controls.Add(Me.txtPromo_src)
        Me.tbListPromo.Controls.Add(Me.Label16)
        Me.tbListPromo.Controls.Add(Me.btnNeu_Promo)
        Me.tbListPromo.Controls.Add(Me.btnBrowProm_src)
        Me.tbListPromo.Controls.Add(Me.Label17)
        Me.tbListPromo.Location = New System.Drawing.Point(4, 23)
        Me.tbListPromo.Name = "tbListPromo"
        Me.tbListPromo.Padding = New System.Windows.Forms.Padding(3)
        Me.tbListPromo.Size = New System.Drawing.Size(403, 162)
        Me.tbListPromo.TabIndex = 0
        Me.tbListPromo.Text = "List Of Promotion"
        '
        'PicBar_Promo
        '
        Me.PicBar_Promo.Image = CType(resources.GetObject("PicBar_Promo.Image"), System.Drawing.Image)
        Me.PicBar_Promo.Location = New System.Drawing.Point(5, 76)
        Me.PicBar_Promo.Name = "PicBar_Promo"
        Me.PicBar_Promo.Size = New System.Drawing.Size(80, 80)
        Me.PicBar_Promo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicBar_Promo.TabIndex = 10
        Me.PicBar_Promo.TabStop = False
        Me.PicBar_Promo.Visible = False
        '
        'btnBrowPromo_dest
        '
        Me.btnBrowPromo_dest.Image = CType(resources.GetObject("btnBrowPromo_dest.Image"), System.Drawing.Image)
        Me.btnBrowPromo_dest.Location = New System.Drawing.Point(310, 48)
        Me.btnBrowPromo_dest.Name = "btnBrowPromo_dest"
        Me.btnBrowPromo_dest.Size = New System.Drawing.Size(26, 23)
        Me.btnBrowPromo_dest.TabIndex = 12
        Me.btnBrowPromo_dest.UseVisualStyleBackColor = True
        '
        'txtProm_dest
        '
        Me.txtProm_dest.Location = New System.Drawing.Point(76, 49)
        Me.txtProm_dest.Name = "txtProm_dest"
        Me.txtProm_dest.Size = New System.Drawing.Size(229, 22)
        Me.txtProm_dest.TabIndex = 11
        '
        'txtPromo_src
        '
        Me.txtPromo_src.Location = New System.Drawing.Point(76, 15)
        Me.txtPromo_src.Name = "txtPromo_src"
        Me.txtPromo_src.Size = New System.Drawing.Size(229, 22)
        Me.txtPromo_src.TabIndex = 7
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(3, 54)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(70, 14)
        Me.Label16.TabIndex = 10
        Me.Label16.Text = "Destination"
        '
        'btnNeu_Promo
        '
        Me.btnNeu_Promo.Enabled = False
        Me.btnNeu_Promo.Image = CType(resources.GetObject("btnNeu_Promo.Image"), System.Drawing.Image)
        Me.btnNeu_Promo.Location = New System.Drawing.Point(343, 12)
        Me.btnNeu_Promo.Name = "btnNeu_Promo"
        Me.btnNeu_Promo.Size = New System.Drawing.Size(54, 59)
        Me.btnNeu_Promo.TabIndex = 9
        Me.btnNeu_Promo.UseVisualStyleBackColor = True
        '
        'btnBrowProm_src
        '
        Me.btnBrowProm_src.Image = CType(resources.GetObject("btnBrowProm_src.Image"), System.Drawing.Image)
        Me.btnBrowProm_src.Location = New System.Drawing.Point(310, 14)
        Me.btnBrowProm_src.Name = "btnBrowProm_src"
        Me.btnBrowProm_src.Size = New System.Drawing.Size(26, 23)
        Me.btnBrowProm_src.TabIndex = 8
        Me.btnBrowProm_src.UseVisualStyleBackColor = True
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(6, 20)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(43, 14)
        Me.Label17.TabIndex = 0
        Me.Label17.Text = "Source"
        '
        'tcListPromo
        '
        Me.tcListPromo.Controls.Add(Me.tbListPromo)
        Me.tcListPromo.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tcListPromo.Location = New System.Drawing.Point(436, 239)
        Me.tcListPromo.Name = "tcListPromo"
        Me.tcListPromo.SelectedIndex = 0
        Me.tcListPromo.Size = New System.Drawing.Size(411, 189)
        Me.tcListPromo.TabIndex = 6
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.BackColor = System.Drawing.Color.Transparent
        Me.Label18.Font = New System.Drawing.Font("Calibri", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(437, 221)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(71, 17)
        Me.Label18.TabIndex = 7
        Me.Label18.Text = "Promotion"
        '
        'bwDSS
        '
        '
        'bwPromo
        '
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(867, 487)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.tcListPromo)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.TabConStock)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TabConSales)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TabCon)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Nirmala UI Semilight", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Main"
        Me.Text = "Report Neutralizer"
        Me.TabCon.ResumeLayout(False)
        Me.TabOutMas.ResumeLayout(False)
        Me.TabOutMas.PerformLayout()
        CType(Me.PicBar, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TPProdMast.ResumeLayout(False)
        Me.TPProdMast.PerformLayout()
        CType(Me.PicBar_ProdMas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabConSales.ResumeLayout(False)
        Me.LP3.ResumeLayout(False)
        Me.LP3.PerformLayout()
        Me.GBLP3.ResumeLayout(False)
        Me.GBLP3.PerformLayout()
        CType(Me.PicBarLP3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TPSalesPerfRep.ResumeLayout(False)
        Me.TPSalesPerfRep.PerformLayout()
        Me.gbRepType_SPR.ResumeLayout(False)
        Me.gbRepType_SPR.PerformLayout()
        CType(Me.PicBar_SPR, System.ComponentModel.ISupportInitialize).EndInit()
        Me.DSStab.ResumeLayout(False)
        Me.DSStab.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.PicBar_DSS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabConStock.ResumeLayout(False)
        Me.TPDistStock.ResumeLayout(False)
        Me.TPDistStock.PerformLayout()
        CType(Me.PicBar_DistStock, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tpDailyStockMutation.ResumeLayout(False)
        Me.tpDailyStockMutation.PerformLayout()
        CType(Me.PicBar_DSM, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbListPromo.ResumeLayout(False)
        Me.tbListPromo.PerformLayout()
        CType(Me.PicBar_Promo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tcListPromo.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TabCon As System.Windows.Forms.TabControl
    Friend WithEvents TabOutMas As System.Windows.Forms.TabPage
    Friend WithEvents btnNeuOutMas As System.Windows.Forms.Button
    Friend WithEvents btnbrow_outmassrc As System.Windows.Forms.Button
    Friend WithEvents txtoutmas_src As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents OFD_src As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnBrow_OutMasDes As System.Windows.Forms.Button
    Friend WithEvents txtDest_OutMas As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents SFD_Dest As System.Windows.Forms.SaveFileDialog
    Friend WithEvents prgbar As System.Windows.Forms.Label
    Friend WithEvents PicBar As System.Windows.Forms.PictureBox
    Friend WithEvents BWOutMas As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TabConSales As System.Windows.Forms.TabControl
    Friend WithEvents LP3 As System.Windows.Forms.TabPage
    Friend WithEvents PicBarLP3 As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_LP3_Dest As System.Windows.Forms.Button
    Friend WithEvents txtLP3_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnNeuLP3 As System.Windows.Forms.Button
    Friend WithEvents btnBrow_LP3_src As System.Windows.Forms.Button
    Friend WithEvents txtLP3_src As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents BWLP3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents GBLP3 As System.Windows.Forms.GroupBox
    Friend WithEvents RBQty As System.Windows.Forms.RadioButton
    Friend WithEvents RBRupiah As System.Windows.Forms.RadioButton
    Friend WithEvents RBAll As System.Windows.Forms.RadioButton
    Friend WithEvents TabConStock As System.Windows.Forms.TabControl
    Friend WithEvents TPDistStock As System.Windows.Forms.TabPage
    Friend WithEvents PicBar_DistStock As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrowDistStock_Dest As System.Windows.Forms.Button
    Friend WithEvents txtDistStock_Dest As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_DistStock As System.Windows.Forms.Button
    Friend WithEvents btnBrowDistStock_src As System.Windows.Forms.Button
    Friend WithEvents txtDistStock_src As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents BWDistStock As System.ComponentModel.BackgroundWorker
    Friend WithEvents TPProdMast As System.Windows.Forms.TabPage
    Friend WithEvents PicBar_ProdMas As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrowProdMas_dest As System.Windows.Forms.Button
    Friend WithEvents txtProdMas_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_ProdMas As System.Windows.Forms.Button
    Friend WithEvents btnBrowProdMas_src As System.Windows.Forms.Button
    Friend WithEvents txtProdMas_src As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents BWProdMas As System.ComponentModel.BackgroundWorker
    Friend WithEvents TPSalesPerfRep As System.Windows.Forms.TabPage
    Friend WithEvents gbRepType_SPR As System.Windows.Forms.GroupBox
    Friend WithEvents RBSPR_Avg As System.Windows.Forms.RadioButton
    Friend WithEvents RBSPR_Val As System.Windows.Forms.RadioButton
    Friend WithEvents RBSPR_Qty As System.Windows.Forms.RadioButton
    Friend WithEvents PicBar_SPR As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_SPR_dest As System.Windows.Forms.Button
    Friend WithEvents txtSPR_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_SPR As System.Windows.Forms.Button
    Friend WithEvents btnBrow_SPR_src As System.Windows.Forms.Button
    Friend WithEvents txtSPR_src As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents BWSalesPerf As System.ComponentModel.BackgroundWorker
    Friend WithEvents tpDailyStockMutation As System.Windows.Forms.TabPage
    Friend WithEvents PicBar_DSM As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrow_DSM_dest As System.Windows.Forms.Button
    Friend WithEvents txtDSM_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_DSM As System.Windows.Forms.Button
    Friend WithEvents btnBrow_DSM_src As System.Windows.Forms.Button
    Friend WithEvents txtDSM_src As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents BWDSM As System.ComponentModel.BackgroundWorker
    Friend WithEvents tbListPromo As System.Windows.Forms.TabPage
    Friend WithEvents PicBar_Promo As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrowPromo_dest As System.Windows.Forms.Button
    Friend WithEvents txtProm_dest As System.Windows.Forms.TextBox
    Friend WithEvents txtPromo_src As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_Promo As System.Windows.Forms.Button
    Friend WithEvents btnBrowProm_src As System.Windows.Forms.Button
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents tcListPromo As System.Windows.Forms.TabControl
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents DSStab As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents rbDSSprod As System.Windows.Forms.RadioButton
    Friend WithEvents rbDSSrupiah As System.Windows.Forms.RadioButton
    Friend WithEvents rbDSSall As System.Windows.Forms.RadioButton
    Friend WithEvents PicBar_DSS As System.Windows.Forms.PictureBox
    Friend WithEvents btnBrowDSS_dest As System.Windows.Forms.Button
    Friend WithEvents txtDSS_dest As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents btnNeu_DSS As System.Windows.Forms.Button
    Friend WithEvents bntBrowDSS_src As System.Windows.Forms.Button
    Friend WithEvents txtDSS_src As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents bwDSS As System.ComponentModel.BackgroundWorker
    Friend WithEvents bwPromo As System.ComponentModel.BackgroundWorker

End Class
