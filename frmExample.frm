VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "�˺� ���ݰ�꼭 SDK ����"
   ClientHeight    =   14070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19275
   LinkTopic       =   "Form1"
   ScaleHeight     =   14070
   ScaleWidth      =   19275
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   13680
      TabIndex        =   118
      Top             =   165
      Width           =   5055
   End
   Begin VB.CommandButton btnCheckID 
      Caption         =   "ID �ߺ� Ȯ��"
      Height          =   410
      Left            =   480
      TabIndex        =   52
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   " �˺� �⺻ API "
      Height          =   3015
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   18495
      Begin VB.Frame Frame15 
         Caption         =   "ȸ������ ����"
         Height          =   2415
         Left            =   9120
         TabIndex        =   109
         Top             =   240
         Width           =   2055
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "ȸ������ ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   111
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "ȸ������ ����"
            Height          =   410
            Left            =   120
            TabIndex        =   110
            Top             =   840
            Width           =   1815
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "��Ʈ�ʰ��� ����Ʈ"
         Height          =   2415
         Index           =   1
         Left            =   6720
         TabIndex        =   106
         Top             =   240
         Width           =   2415
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "��Ʈ�� �ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   108
            Top             =   360
            Width           =   2150
         End
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   107
            Top             =   840
            Width           =   2150
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "�������� ����Ʈ"
         Height          =   2415
         Index           =   0
         Left            =   4440
         TabIndex        =   62
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton btnGetUseHistoryURL 
            Caption         =   "����Ʈ ��볻�� URL"
            Height          =   375
            Left            =   120
            TabIndex        =   114
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton btnGetPaymentURL 
            Caption         =   "����Ʈ �������� URL"
            Height          =   375
            Left            =   120
            TabIndex        =   113
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   " ����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   63
            Top             =   840
            Width           =   1935
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " ���������� ����"
         Height          =   2415
         Left            =   11280
         TabIndex        =   12
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton btnGetTaxCertInfo 
            Caption         =   "������ ���� Ȯ��"
            Height          =   375
            Left            =   120
            TabIndex        =   119
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton btnCheckCertValidation 
            Caption         =   "������ ��ȿ�� Ȯ��"
            Height          =   375
            Left            =   120
            TabIndex        =   67
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton btnGetTaxCertURL 
            Caption         =   " ������ ��� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   60
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton btnCertificateExpireDate 
            Caption         =   "������ ������ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " ȸ������"
         Height          =   2415
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1695
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "ȸ�� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " ����Ʈ ����"
         Height          =   2415
         Left            =   1920
         TabIndex        =   8
         Top             =   240
         Width           =   2415
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "�������� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   59
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "��� �ܰ� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "����� ����"
         Height          =   2415
         Left            =   13560
         TabIndex        =   7
         Top             =   240
         Width           =   2055
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "����� ���� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   116
            Top             =   1800
            Width           =   1815
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "����� ��� ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   115
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CommandButton btnGetContactInfo 
            Caption         =   "����� ���� Ȯ��"
            Height          =   375
            Left            =   120
            TabIndex        =   112
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "����� �߰�"
            Height          =   410
            Left            =   120
            TabIndex        =   53
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " �˺� �⺻ URL"
         Height          =   2415
         Left            =   15720
         TabIndex        =   5
         Top             =   240
         Width           =   2655
         Begin VB.CommandButton btnGetSealURL 
            Caption         =   "�ΰ� �� ÷�ι��� ��� URL"
            Height          =   375
            Left            =   120
            TabIndex        =   61
            Top             =   840
            Width           =   2415
         End
         Begin VB.CommandButton btnGetAccessURL 
            Caption         =   " �˺� �α��� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   2415
         End
      End
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   6000
      TabIndex        =   3
      Text            =   "testkorea"
      Top             =   165
      Width           =   1935
   End
   Begin VB.TextBox txtCorpNum 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Text            =   "1234567890"
      Top             =   180
      Width           =   1935
   End
   Begin VB.Frame Frame7 
      Caption         =   " ���ݰ�꼭 ���� ���"
      Height          =   10185
      Left            =   240
      TabIndex        =   13
      Top             =   3720
      Width           =   18495
      Begin VB.Frame Frame16 
         Caption         =   " (����) ��ù��� ���μ���"
         Height          =   3255
         Left            =   720
         TabIndex        =   54
         Top             =   1200
         Width           =   3255
         Begin VB.CommandButton btnRegistIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "��ù���"
            Height          =   495
            Left            =   360
            Style           =   1  '�׷���
            TabIndex        =   73
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton btnCancelIssue_sub 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�������"
            Height          =   495
            Left            =   360
            Style           =   1  '�׷���
            TabIndex        =   56
            Top             =   2110
            Width           =   975
         End
         Begin VB.CommandButton btnDelete_sub 
            Caption         =   "����"
            Height          =   495
            Left            =   1920
            Style           =   1  '�׷���
            TabIndex        =   55
            Top             =   2110
            Width           =   975
         End
         Begin VB.Line Line19 
            X1              =   840
            X2              =   2475
            Y1              =   2355
            Y2              =   2355
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            FillColor       =   &H00E0E0E0&
            Height          =   780
            Left            =   240
            Top             =   360
            Width           =   2625
         End
         Begin VB.Line Line18 
            X1              =   840
            X2              =   840
            Y1              =   2400
            Y2              =   960
         End
      End
      Begin VB.CommandButton btnGetEmailPublicKeys 
         Caption         =   "�������ڸ��� ���"
         Height          =   375
         Left            =   16080
         TabIndex        =   50
         Top             =   240
         Width           =   1965
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   4800
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame14 
         Caption         =   " ����/�μ�"
         Height          =   3960
         Left            =   10080
         TabIndex        =   45
         Top             =   6120
         Width           =   5490
         Begin VB.CommandButton btnGetOldPrintURL 
            Caption         =   "(��)�μ� �˾� URL"
            Height          =   375
            Left            =   210
            TabIndex        =   76
            Top             =   1800
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPDFURL 
            Caption         =   "PDF �ٿ�ε� URL"
            Height          =   375
            Left            =   3120
            TabIndex        =   75
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnGetViewURL 
            Caption         =   "���ݰ�꼭 �˾� URL (�޴�x)"
            Height          =   390
            Left            =   210
            TabIndex        =   74
            Top             =   840
            Width           =   2745
         End
         Begin VB.CommandButton btnGetEPrintUrl 
            Caption         =   "���޹޴��� �μ� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   51
            Top             =   2280
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPopUpURL 
            Caption         =   "���ݰ�꼭 �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   49
            Top             =   390
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPrintURL 
            Caption         =   "�μ� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   48
            Top             =   1305
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMassPrintURL 
            Caption         =   "�뷮 �μ� �˾� URL"
            Height          =   390
            Left            =   3120
            TabIndex        =   47
            Top             =   840
            Width           =   2175
         End
         Begin VB.CommandButton btnGetMailURL 
            Caption         =   "���ݰ�꼭 ���ϸ�ũ URL"
            Height          =   390
            Left            =   210
            TabIndex        =   46
            Top             =   2760
            Width           =   2745
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   " ��Ÿ URL "
         Height          =   3975
         Left            =   15600
         TabIndex        =   40
         Top             =   6120
         Width           =   2385
         Begin VB.CommandButton btnGetURL_PWBOX 
            Caption         =   "���� ���� �����"
            Height          =   390
            Index           =   2
            Left            =   210
            TabIndex        =   122
            Top             =   1320
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_SWBOX 
            Caption         =   "���� ���� �����"
            Height          =   390
            Index           =   1
            Left            =   210
            TabIndex        =   121
            Top             =   840
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_TBOX 
            Caption         =   "�ӽ� ������"
            Height          =   390
            Index           =   0
            Left            =   210
            TabIndex        =   44
            Top             =   360
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_SBOX 
            Caption         =   "���� ������"
            Height          =   390
            Left            =   210
            TabIndex        =   43
            Top             =   1785
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_PBOX 
            Caption         =   "���� ������"
            Height          =   390
            Left            =   210
            TabIndex        =   42
            Top             =   2220
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_WRITE 
            Caption         =   "���� �����ۼ�"
            Height          =   390
            Left            =   210
            TabIndex        =   41
            Top             =   2640
            Width           =   1845
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   " �ΰ� ���"
         Height          =   3975
         Left            =   5400
         TabIndex        =   38
         Top             =   6120
         Width           =   4545
         Begin VB.CommandButton btnAssignmgtkey 
            Caption         =   "������ȣ �Ҵ�"
            Height          =   390
            Left            =   240
            TabIndex        =   85
            Top             =   1800
            Width           =   1965
         End
         Begin VB.CommandButton btnListemailconfig 
            Caption         =   "�˸����� ���۸�� ��ȸ"
            Height          =   390
            Left            =   2280
            TabIndex        =   84
            Top             =   1320
            Width           =   2085
         End
         Begin VB.CommandButton btnUpdateemailconfig 
            Caption         =   "�˸����� ���ۼ��� ����"
            Height          =   390
            Left            =   2280
            TabIndex        =   83
            Top             =   1800
            Width           =   2085
         End
         Begin VB.CommandButton btnGetSendToNTSConfig 
            Caption         =   "����û ���� ���� Ȯ��"
            Height          =   390
            Left            =   2280
            TabIndex        =   82
            Top             =   2280
            Width           =   2085
         End
         Begin VB.CommandButton btnSendFAX 
            Caption         =   "�ѽ� ����"
            Height          =   375
            Left            =   240
            TabIndex        =   66
            Top             =   1320
            Width           =   1965
         End
         Begin VB.CommandButton btnSendSMS 
            Caption         =   "���� ����"
            Height          =   375
            Left            =   240
            TabIndex        =   65
            Top             =   840
            Width           =   1965
         End
         Begin VB.CommandButton btnDetachStatement 
            Caption         =   "���ڸ��� ÷������"
            Height          =   390
            Left            =   2280
            TabIndex        =   58
            Top             =   840
            Width           =   2085
         End
         Begin VB.CommandButton btnAttachStatement 
            Caption         =   "���ڸ��� ÷��"
            Height          =   390
            Left            =   2280
            TabIndex        =   57
            Top             =   360
            Width           =   2085
         End
         Begin VB.CommandButton btnSendEmail 
            Caption         =   "�̸��� ����"
            Height          =   390
            Left            =   240
            TabIndex        =   39
            Top             =   390
            Width           =   1965
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   " ���� Ȯ��"
         Height          =   3975
         Left            =   2640
         TabIndex        =   33
         Top             =   6120
         Width           =   2505
         Begin VB.CommandButton btnGetXML 
            Caption         =   "������ Ȯ�� - XML"
            Height          =   390
            Left            =   195
            TabIndex        =   120
            Top             =   1800
            Width           =   2085
         End
         Begin VB.CommandButton btnSearch 
            Caption         =   "��� ��ȸ"
            Height          =   390
            Left            =   195
            TabIndex        =   86
            Top             =   2280
            Width           =   2085
         End
         Begin VB.CommandButton btnGetBulkResult 
            Caption         =   "�ʴ뷮 ������� Ȯ��"
            Height          =   390
            Left            =   195
            TabIndex        =   81
            Top             =   3240
            Width           =   2085
         End
         Begin VB.CommandButton btnGetDetailInfo 
            Caption         =   "������ Ȯ��"
            Height          =   390
            Left            =   195
            TabIndex        =   37
            Top             =   1320
            Width           =   2085
         End
         Begin VB.CommandButton btnGetLogs 
            Caption         =   "���� �����̷�"
            Height          =   390
            Left            =   195
            TabIndex        =   36
            Top             =   2760
            Width           =   2085
         End
         Begin VB.CommandButton btnGetInfos 
            Caption         =   "���� �뷮 Ȯ��"
            Height          =   390
            Left            =   210
            TabIndex        =   35
            Top             =   825
            Width           =   2085
         End
         Begin VB.CommandButton btnGetInfo 
            Caption         =   "���� Ȯ��"
            Height          =   390
            Left            =   210
            TabIndex        =   34
            Top             =   390
            Width           =   2085
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   " ÷������ "
         Height          =   3960
         Left            =   120
         TabIndex        =   28
         Top             =   6135
         Width           =   2265
         Begin VB.CommandButton btnDeleteFile 
            Caption         =   "���� ����"
            Height          =   390
            Left            =   210
            TabIndex        =   32
            Top             =   1650
            Width           =   1845
         End
         Begin VB.TextBox txtFileID 
            Height          =   330
            Left            =   210
            TabIndex        =   31
            Text            =   "���Ͼ��̵�"
            Top             =   1245
            Width           =   1845
         End
         Begin VB.CommandButton btnGetFiles 
            Caption         =   "÷�� ���"
            Height          =   390
            Left            =   210
            TabIndex        =   30
            Top             =   795
            Width           =   1845
         End
         Begin VB.CommandButton btnAttachFile 
            Caption         =   "���� ÷��"
            Height          =   390
            Left            =   210
            TabIndex        =   29
            Top             =   345
            Width           =   1845
         End
      End
      Begin VB.CommandButton btnSendToNTS 
         BackColor       =   &H00C0C0FF&
         Caption         =   "����û ��� ����"
         Height          =   375
         Left            =   2280
         Style           =   1  '�׷���
         TabIndex        =   27
         Top             =   4560
         Width           =   4200
      End
      Begin VB.Frame Frame8 
         Caption         =   " �ӽ����� ���� ���μ���"
         Height          =   3255
         Left            =   4200
         TabIndex        =   20
         Top             =   1200
         Width           =   4695
         Begin VB.CommandButton btnRegister 
            BackColor       =   &H00C0C0FF&
            Caption         =   "���"
            Height          =   375
            Left            =   1305
            Style           =   1  '�׷���
            TabIndex        =   25
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton btnUpdate 
            BackColor       =   &H00C0C0FF&
            Caption         =   "����"
            Height          =   375
            Left            =   2265
            Style           =   1  '�׷���
            TabIndex        =   24
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton btnDelete 
            Caption         =   "����"
            Height          =   375
            Left            =   3345
            Style           =   1  '�׷���
            TabIndex        =   23
            Top             =   2280
            Width           =   975
         End
         Begin VB.CommandButton btnCancelIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�������"
            Height          =   375
            Left            =   360
            Style           =   1  '�׷���
            TabIndex        =   22
            Top             =   2280
            Width           =   975
         End
         Begin VB.CommandButton btnIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "����"
            Height          =   495
            Left            =   360
            Style           =   1  '�׷���
            TabIndex        =   21
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "�ӽ�����"
            Height          =   180
            Left            =   465
            TabIndex        =   26
            Top             =   555
            Width           =   720
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            FillColor       =   &H00E0E0E0&
            Height          =   660
            Left            =   300
            Top             =   345
            Width           =   4200
         End
         Begin VB.Line Line3 
            X1              =   3840
            X2              =   3840
            Y1              =   2550
            Y2              =   780
         End
         Begin VB.Line Line2 
            X1              =   900
            X2              =   4200
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Line Line1 
            X1              =   840
            X2              =   840
            Y1              =   2500
            Y2              =   680
         End
      End
      Begin VB.ComboBox cboMgtKeyType 
         Height          =   300
         Left            =   2520
         TabIndex        =   19
         Text            =   "SELL"
         Top             =   300
         Width           =   1335
      End
      Begin VB.CommandButton checkMgtKeyInUse 
         Caption         =   "������ȣ ��뿩�� Ȯ��"
         Height          =   375
         Left            =   6840
         TabIndex        =   18
         Top             =   240
         Width           =   2190
      End
      Begin VB.TextBox txtMgtKey 
         Height          =   330
         Left            =   3960
         TabIndex        =   17
         Top             =   285
         Width           =   2775
      End
      Begin VB.Frame Frame21 
         Height          =   615
         Left            =   9240
         TabIndex        =   70
         Top             =   120
         Width           =   3615
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   ": ���޹޴��� ó��"
            Height          =   180
            Left            =   2040
            TabIndex        =   72
            Top             =   270
            Width           =   1440
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   ": ������ ó��"
            Height          =   180
            Left            =   480
            TabIndex        =   71
            Top             =   270
            Width           =   1080
         End
         Begin VB.Shape Shape8 
            BackColor       =   &H00C0C0FF&
            BorderColor     =   &H00404040&
            FillColor       =   &H00C0C0FF&
            FillStyle       =   0  '�ܻ�
            Height          =   255
            Left            =   120
            Top             =   240
            Width           =   255
         End
         Begin VB.Shape Shape7 
            BackColor       =   &H00FFFFC0&
            BorderColor     =   &H00404040&
            FillColor       =   &H00FFFFC0&
            FillStyle       =   0  '�ܻ�
            Height          =   255
            Left            =   1680
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   " ������ ���μ���"
         Height          =   4215
         Left            =   480
         TabIndex        =   68
         Top             =   840
         Width           =   8775
      End
      Begin VB.Frame Frame22 
         Caption         =   "�ʴ뷮 ����"
         Height          =   735
         Left            =   480
         TabIndex        =   77
         Top             =   5160
         Width           =   8775
         Begin VB.CommandButton btnBulkSubmit 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�ʴ뷮 ���� ����"
            Height          =   375
            Left            =   6960
            TabIndex        =   80
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtSubmitID 
            Height          =   330
            Left            =   2400
            TabIndex        =   79
            Top             =   290
            Width           =   4455
         End
         Begin VB.Label Label8 
            Caption         =   "���� ���̵�(SubmitID) : "
            Height          =   255
            Left            =   240
            TabIndex        =   78
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   " ������ ���μ���"
         Height          =   4215
         Left            =   9480
         TabIndex        =   87
         Top             =   840
         Width           =   8415
         Begin VB.Frame Frame9 
            Caption         =   " �ӽ����� ������ ���μ��� "
            Height          =   3255
            Left            =   4200
            TabIndex        =   96
            Top             =   360
            Width           =   4095
            Begin VB.CommandButton btnRegister_rev 
               BackColor       =   &H00FFFFC0&
               Caption         =   "���"
               Height          =   375
               Left            =   1515
               Style           =   1  '�׷���
               TabIndex        =   104
               Top             =   465
               Width           =   855
            End
            Begin VB.CommandButton btnUpdate_rev 
               BackColor       =   &H00FFFFC0&
               Caption         =   "����"
               Height          =   375
               Left            =   2475
               Style           =   1  '�׷���
               TabIndex        =   103
               Top             =   465
               Width           =   855
            End
            Begin VB.CommandButton btnRequest 
               BackColor       =   &H00FFFFC0&
               Caption         =   "��)�����û"
               Height          =   420
               Left            =   320
               Style           =   1  '�׷���
               TabIndex        =   102
               Top             =   1200
               Width           =   1920
            End
            Begin VB.CommandButton btnIssue_rev 
               BackColor       =   &H00C0C0FF&
               Caption         =   "����"
               Height          =   420
               Left            =   360
               Style           =   1  '�׷���
               TabIndex        =   101
               Top             =   1800
               Width           =   855
            End
            Begin VB.CommandButton btnDelete_rev 
               Caption         =   "����"
               Height          =   420
               Left            =   2760
               Style           =   1  '�׷���
               TabIndex        =   100
               Top             =   2520
               Width           =   855
            End
            Begin VB.CommandButton btnCancelIssue_rev 
               BackColor       =   &H00C0C0FF&
               Caption         =   "�������"
               Height          =   420
               Left            =   360
               Style           =   1  '�׷���
               TabIndex        =   99
               Top             =   2520
               Width           =   855
            End
            Begin VB.CommandButton btnRequestCancel 
               BackColor       =   &H00FFFFC0&
               Caption         =   "��û���"
               Height          =   420
               Left            =   2760
               Style           =   1  '�׷���
               TabIndex        =   98
               Top             =   1200
               Width           =   855
            End
            Begin VB.CommandButton btnRefuse 
               BackColor       =   &H00C0C0FF&
               Caption         =   "�ź�"
               Height          =   420
               Left            =   1320
               Style           =   1  '�׷���
               TabIndex        =   97
               Top             =   1800
               Width           =   855
            End
            Begin VB.Line Line17 
               X1              =   3240
               X2              =   3240
               Y1              =   2630
               Y2              =   1500
            End
            Begin VB.Line Line13 
               X1              =   750
               X2              =   750
               Y1              =   2685
               Y2              =   840
            End
            Begin VB.Shape Shape3 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  '�������� ����
               FillColor       =   &H00E0E0E0&
               Height          =   660
               Left            =   240
               Top             =   315
               Width           =   3360
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  '����
               Caption         =   "�ӽ�����"
               Height          =   180
               Left            =   675
               TabIndex        =   105
               Top             =   540
               Width           =   720
            End
            Begin VB.Line Line14 
               X1              =   1080
               X2              =   2925
               Y1              =   2760
               Y2              =   2760
            End
            Begin VB.Line Line16 
               X1              =   1680
               X2              =   1680
               Y1              =   1560
               Y2              =   2760
            End
            Begin VB.Line Line25 
               X1              =   2040
               X2              =   2880
               Y1              =   1440
               Y2              =   1440
            End
         End
         Begin VB.Frame Frame18 
            Caption         =   " (����) ��ÿ�û ���μ���"
            Height          =   3255
            Left            =   360
            TabIndex        =   88
            Top             =   360
            Width           =   3615
            Begin VB.CommandButton btnRefuse_sub 
               BackColor       =   &H00C0C0FF&
               Caption         =   "�ź�"
               Height          =   420
               Left            =   1440
               Style           =   1  '�׷���
               TabIndex        =   94
               Top             =   1560
               Width           =   855
            End
            Begin VB.CommandButton btnRequestCancel_sub 
               BackColor       =   &H00FFFFC0&
               Caption         =   "��û���"
               Height          =   420
               Left            =   2520
               Style           =   1  '�׷���
               TabIndex        =   93
               Top             =   1560
               Width           =   855
            End
            Begin VB.CommandButton btnCancelIssue_rev_sub 
               BackColor       =   &H00C0C0FF&
               Caption         =   "�������"
               Height          =   420
               Left            =   330
               Style           =   1  '�׷���
               TabIndex        =   92
               Top             =   2520
               Width           =   855
            End
            Begin VB.CommandButton btnDelete_rev_sub 
               Caption         =   "����"
               Height          =   420
               Left            =   2520
               Style           =   1  '�׷���
               TabIndex        =   91
               Top             =   2520
               Width           =   855
            End
            Begin VB.CommandButton btnIssue_rev_sub 
               BackColor       =   &H00C0C0FF&
               Caption         =   "����"
               Height          =   420
               Left            =   330
               Style           =   1  '�׷���
               TabIndex        =   90
               Top             =   1560
               Width           =   855
            End
            Begin VB.CommandButton btnRegistRequest 
               BackColor       =   &H00FFFFC0&
               Caption         =   "��ÿ�û"
               Height          =   420
               Left            =   1560
               Style           =   1  '�׷���
               TabIndex        =   89
               Top             =   480
               Width           =   1455
            End
            Begin VB.Line Line21 
               X1              =   2880
               X2              =   2880
               Y1              =   960
               Y2              =   2760
            End
            Begin VB.Line Line23 
               X1              =   1320
               X2              =   1320
               Y1              =   960
               Y2              =   1200
            End
            Begin VB.Shape Shape6 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  '�������� ����
               FillColor       =   &H00E0E0E0&
               Height          =   660
               Left            =   120
               Top             =   360
               Width           =   3360
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  '����
               Caption         =   "��ÿ�û"
               Height          =   180
               Left            =   480
               TabIndex        =   95
               Top             =   600
               Width           =   720
            End
            Begin VB.Line Line15 
               X1              =   720
               X2              =   720
               Y1              =   1200
               Y2              =   2760
            End
            Begin VB.Line Line20 
               X1              =   1200
               X2              =   2760
               Y1              =   2760
               Y2              =   2760
            End
            Begin VB.Line Line22 
               X1              =   1905
               X2              =   1905
               Y1              =   1200
               Y2              =   2760
            End
            Begin VB.Line Line24 
               X1              =   720
               X2              =   1900
               Y1              =   1200
               Y2              =   1200
            End
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "������ȣ( MgtKey) : "
         Height          =   180
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   2115
      End
   End
   Begin VB.Label URL 
      AutoSize        =   -1  'True
      Caption         =   "URL : "
      Height          =   180
      Left            =   13080
      TabIndex        =   117
      Top             =   240
      Width           =   525
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�ӽ�����"
      Height          =   180
      Left            =   8040
      TabIndex        =   69
      Top             =   5400
      Width           =   720
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  '�������� ����
      FillColor       =   &H00E0E0E0&
      Height          =   660
      Left            =   10200
      Top             =   4440
      Width           =   3360
   End
   Begin VB.Label Label2 
      Caption         =   "�˺�ȸ�� ���̵� : "
      Height          =   180
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "�˺�ȸ�� ����ڹ�ȣ :"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1920
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' �˺� ���ڼ��ݰ�꼭 API VB 6.0 SDK Example
'
' - ������Ʈ ���� : 2022-11-01
' - ���� ������� ����ó : 1600-9854
' - ���� ������� �̸��� : code@linkhubcorp.com
' - VB6 SDK ������ �ȳ� : https://developers.popbill.com/guide/taxinvoice/vb/getting-started/tutorial

' <�׽�Ʈ �������� �غ����>
' 1) 30, 33�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
' 2) ���ڼ��ݰ�꼭 ������ ���� ������������ ����մϴ�.
'    - �˺�����Ʈ �α��� > [���ڼ��ݰ�꼭] > [ȯ�漳��]
'      > [���������� ����]
'    - ���������� ��� �˾� URL (GetTaxCertURL API)�� �̿��Ͽ� ���
'
'=========================================================================

Option Explicit

'��ũ���̵�
Private Const LinkID = "TESTER"

'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'���ݰ�꼭 ��ü ����
Private TaxinvoiceService As New PBTIService


'=========================================================================
' ����ڹ�ȣ�� ��ȸ�Ͽ� ����ȸ�� ���Կ��θ� Ȯ���մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/member#CheckIsMember
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.CheckIsMember(txtCorpNum.Text, LinkID)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ϰ��� �ϴ� ���̵��� �ߺ����θ� Ȯ���մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/member#CheckID
'=========================================================================
Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ڼ��ݰ�꼭 PDF ������ �ٿ� ���� �� �ִ� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ������å�� 30���� ��ȿ�ð��� ������, ��ȿ�ð� ���� ȣ��� ���������� �������� ȣ����� �ʽ��ϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/view#GetPDFURL
'=========================================================================
Private Sub btnGetPDFURL_Click()
    Dim URL As String
    Dim KeyType As MgtKeyType
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    URL = TaxinvoiceService.GetPDFURL(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    txtUserID.Text = URL
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' �˺� ���������� ��ϵ� ������������ ������ Ȯ���մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/cert#GetTaxCertInfo
'=========================================================================
Private Sub btnGetTaxCertInfo_Click()
    Dim CertInfo As PBTaxinvoiceCertificate
    Dim tmp As String
    
    Set CertInfo = TaxinvoiceService.GetTaxCertInfo(txtCorpNum.Text)
    
    If CertInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "regDT (����Ͻ�) : " + CertInfo.regDT + vbCrLf
    tmp = tmp + "expireDT (�����Ͻ�) : " + CertInfo.expireDT + vbCrLf
    tmp = tmp + "issuerDN (������ �߱��� DN) : " + CertInfo.issuerDN + vbCrLf
    tmp = tmp + "subjectDN (��ϵ� ������ DN) : " + CertInfo.subjectDN + vbCrLf
    tmp = tmp + "issuerName (������ ����) : " + CertInfo.issuerName + vbCrLf
    tmp = tmp + "oid (OID) : " + CertInfo.oid + vbCrLf
    tmp = tmp + "regContactName (��� ����� ����) : " + CertInfo.regContactName + vbCrLf
    tmp = tmp + "regContactID (��� ����� ���̵�  ) : " + CertInfo.regContactID + vbCrLf
    
    MsgBox tmp

End Sub


'=========================================================================
' ���ݰ�꼭 1���� ������ ������(����Ʈ ���, ���� �޴� �� ��ư ����)�� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/view#GetViewURL
'=========================================================================
Private Sub btnGetViewURL_Click()
    Dim URL As String
    Dim KeyType As MgtKeyType
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    URL = TaxinvoiceService.GetViewURL(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ����ڸ� ����ȸ������ ����ó���մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/member#JoinMember
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '���̵�, 6���̻� 50�� �̸�
    joinData.id = "testkorea01"
    
    '��й�ȣ, 8�� �̻� 20�� ����(����, ����, Ư������ ����)
    joinData.Password = "asdf$%^123"
    
    '��Ʈ�ʸ�ũ ���̵�
    joinData.LinkID = LinkID
    
    '����ڹ�ȣ, '-'����, 10�ڸ�
    joinData.CorpNum = "1234567890"
    
    '��ǥ�ڼ���, �ִ� 100��
    joinData.ceoname = "��ǥ�ڼ���"
    
    '��ȣ��, �ִ� 200��
    joinData.corpName = "ȸ����ȣ"
    
    '����� �ּ�, �ִ� 300��
    joinData.addr = "�ּ�"
    
    '����, �ִ� 100��
    joinData.bizType = "����"
    
    '����, �ִ� 100��
    joinData.bizClass = "����"

    '����� ����, �ִ� 100��
    joinData.ContactName = "����ڼ���"
    
    '����� �̸���, �ִ� 100��
    joinData.ContactEmail = "test@test.com"
    
    '����� ����ó, �ִ� 20��
    joinData.ContactTEL = "02-999-9999"
    
    
    Set Response = TaxinvoiceService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ݰ�꼭 ����� ���ݵǴ� ����Ʈ �ܰ��� Ȯ���մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/point#GetUnitCost
'=========================================================================
Private Sub btnUnitCost_Click()
    Dim unitCost As Double
    
    unitCost = TaxinvoiceService.GetUnitCost(txtCorpNum.Text)
    
    If unitCost < 0 Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "����ܰ� : " + CStr(unitCost)
End Sub

'=========================================================================
' �˺� ���ڼ��ݰ�꼭 API ���� ���������� Ȯ���մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/point#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = TaxinvoiceService.GetChargeInfo(txtCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (����ܰ�) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' �˺� ���������� ��ϵ� �������� �������� Ȯ���մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/cert#GetCertificateExpireDate
'=========================================================================
Private Sub btnCertificateExpireDate_Click()
    Dim expireDate As String
    
    expireDate = TaxinvoiceService.GetCertificateExpireDate(txtCorpNum.Text)
    
    If expireDate = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������������ : " + expireDate
End Sub

'=========================================================================
' ���ڼ��ݰ�꼭 ���࿡ �ʿ��� �������� �˺� ���������� ����ϱ� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/cert#GetTaxCertURL
'=========================================================================
Private Sub btnGetTaxCertURL_Click()
    Dim URL As String
           
    URL = TaxinvoiceService.GetTaxCertURL(txtCorpNum.Text, txtUserID.Text)

    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
    'Internet Explorer Browser ȣ��
    Dim IE As Object
    Dim strResult As String
    Dim strSiteName As String
   
    Set IE = CreateObject("InternetExplorer.Application")
    strSiteName = URL
    IE.Navigate strSiteName
    With IE
        .Resizable = True
        .MenuBar = True
        .Toolbar = True
        .AddressBar = True
        .Visible = True
        .StatusBar = True
        .Left = 0
        .Top = 0
        .Height = 800
        .Width = 800
        .StatusText = "�˺� ���������� ��� URL"
    End With
    
    Set IE = Nothing
End Sub

'=========================================================================
' �˺� ���������� ��ϵ� �������� ��ȿ���� Ȯ���մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/cert#CheckCertValidation
'=========================================================================
Private Sub btnCheckCertValidation_Click()
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.CheckCertValidation(txtCorpNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �˺� ����Ʈ�� �α��� ���·� ������ �� �ִ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/member#GetAccessURL
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim URL As String
           
    URL = TaxinvoiceService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ���ݰ�꼭�� ÷���� �ΰ�, ����ڵ����, ����纻�� ����ϴ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/etc#GetSealURL
'=========================================================================
Private Sub btnGetSealURL_Click()
    Dim URL As String
           
    URL = TaxinvoiceService.GetSealURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
   
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� �����(�˺� �α��� ����)�� �߰��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/member#RegistContact
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�, 6�� �̻� 50�� �̸�
    joinData.id = "contactID04"
    
    '��й�ȣ, 8�� �̻� 20�� ����(����, ����, Ư������ ����)
    joinData.Password = "asdf#$%123"
    
    '����ڸ�, �ִ� 100��
    joinData.personName = "����ڸ�"
    
    '����� ����ó, �ִ� 20��
    joinData.tel = "070-1234-1234"
    
    '����� �����ּ�, �ִ� 100��
    joinData.email = "test@test.com"
    
    '����� ����, 1-���� / 2-�б� / 3-ȸ��
    joinData.searchRole = 3
    
        
    Set Response = TaxinvoiceService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ������ Ȯ���մϴ�.
' https://developers.popbill.com/reference/taxinvoice/vb/api/member#GetContactInfo
'=========================================================================
Private Sub btnGetContactInfo_Click()
    Dim tmp As String
    Dim info As PBContactInfo
    Dim ContactID As String
    
    'Ȯ���� ����� ���̵�
    ContactID = "testkorea"
    
    Set info = TaxinvoiceService.GetContactInfo(txtCorpNum.Text, ContactID)
    
    If info Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(���̵�) | personName(����) | email(�̸���) | tel(����ó) | " _
         + "regDT(����Ͻ�) | searchRole(����� ����) | mgrYN(������ ����) | state(����) " + vbCrLf
    
   
    tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.tel + " | " _
            + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
        
    MsgBox tmp
End Sub

'=========================================================================
'  ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ����� Ȯ���մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/member#ListContact
'=========================================================================
Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = TaxinvoiceService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(���̵�) | personName(����) | email(�̸���) | tel(����ó) | " _
         + "regDT(����Ͻ�) | searchRole(����� ����) | mgrYN(������ ����) | state(����) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ������ �����մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/member#UpdateContact
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�
    joinData.id = "RegistNodeTest03"
    
    '����� ����, �ִ� 100��
    joinData.personName = "VB6 ����ڸ�"
    
    '����� ����ó, �ִ� 20��
    joinData.tel = "070-1234-1234"
    
    '����� �̸���, �ִ� 100��
    joinData.email = "test@test.com"

    '����� ����, 1-���� / 2-�б� / 3-ȸ��
    joinData.searchRole = 3
                
    Set Response = TaxinvoiceService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� Ȯ���մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/member#GetCorpInfo
'=========================================================================
Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = TaxinvoiceService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname (��ǥ�ڸ�) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName (��ȣ) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr (�ּ�) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType (����) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass (����) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� �����մϴ�
'- https://developers.popbill.com/reference/taxinvoice/vb/api/member#UpdateCorpInfo
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '��ǥ�ڸ�, �ִ� 100��
    CorpInfo.ceoname = "��ǥ��"
    
    '��ȣ, �ִ� 200��
    CorpInfo.corpName = "��ȣ"
    
    '�ּ�, �ִ� 300��
    CorpInfo.addr = "����Ư����"
    
    '����, �ִ� 100��
    CorpInfo.bizType = "����"
    
    '����, �ִ� 100��
    CorpInfo.bizClass = "����"
    
    Set Response = TaxinvoiceService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/point#GetBalance
'=========================================================================
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = TaxinvoiceService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "����ȸ�� �ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ������ ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/point#GetChargeURL
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim URL As String
           
    URL = TaxinvoiceService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ �������� Ȯ���� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/point#GetPaymentURL
'=========================================================================
Private Sub btnGetPaymentURL_Click()
    Dim URL As String
           
    URL = TaxinvoiceService.GetPaymentURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ��볻�� Ȯ���� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/point#GetUseHistoryURL
'=========================================================================
Private Sub btnGetUseHistoryURL_Click()
    Dim URL As String
           
    URL = TaxinvoiceService.GetUseHistoryURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ��Ʈ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/point#GetPartnerBalance
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = TaxinvoiceService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "��Ʈ�� �ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ��Ʈ�� ����Ʈ ������ ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/point#GetPartnerURL
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim URL As String
           
    URL = TaxinvoiceService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ��Ʈ�ʰ� ���ݰ�꼭 ���� �������� �Ҵ��ϴ� ������ȣ�� ��뿩�θ� Ȯ���մϴ�.
' - �̹� ��� ���� ������ȣ�� �ߺ� ����� �Ұ��ϰ�, ���ݰ�꼭�� ������ ��쿡�� ������ȣ�� ������ �����մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/info#CheckMgtKeyInUse
'=========================================================================
Private Sub checkMgtKeyInUse_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.checkMgtKeyInUse(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ڼ��ݰ�꼭 ���������� ���� ����� Ȯ���մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/etc#GetEmailPublicKeys
'=========================================================================
Private Sub btnGetEmailPublicKeys_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim email As Variant
    
    Set resultList = TaxinvoiceService.GetEmailPublicKeys(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "�������� �̸��� ���" + vbCrLf
    For Each email In resultList
        tmp = tmp + email + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' �ۼ��� ���ݰ�꼭 �����͸� �˺��� ����� ���ÿ� ����(���ڼ���)�Ͽ� "����Ϸ�" ���·� ó���մϴ�.
' - ���ݰ�꼭 ����û ���� ��å [https://developers.popbill.com/guide/taxinvoice/vb/introduction/policy-of-send-to-nts]
' - "����Ϸ�"�� ���ڼ��ݰ�꼭�� ����û ���� ������ �������(CancelIssue API) �Լ��� ����û �Ű� ��󿡼� ������ �� �ֽ��ϴ�.
' - �ӽ�����(Register API) �Լ��� ����(Issue API) �Լ��� �� ���� ���μ����� ó���մϴ�.
' - ���ݰ�꼭 ������ ���ؼ� �������� �������� �˺� ���������� ������� �Ǿ�� �մϴ�.
'   �� ����Ź������ ���, ��Ź���� ������ ����� �ʿ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#RegistIssue
'=========================================================================
Private Sub btnRegistIssue_Click()
    Dim Taxinvoice As New PBTaxinvoice
    
    '[�ʼ�] �ۼ�����, ǥ������ (yyyyMMdd)
    Taxinvoice.writeDate = "20220101"
    
    '[�ʼ�] ��������, [������, ������, ����Ź] �� ����
    Taxinvoice.issueType = "������"
    
    '[�ʼ�] {������, ������} �� ����, '������'�� ������ ���μ��������� �̿밡��
    '- ������(������ ����), ������(���޹޴��� ����)
    Taxinvoice.chargeDirection = "������"
    
    '[�ʼ�] ����/û��, [����, û��, ����] �� ����
    Taxinvoice.purposeType = "����"
        
    '[�ʼ�] ��������, [����, ����, �鼼] �� ����
    Taxinvoice.taxType = "����"
    
    
    '=========================================================================
    '                              ������ ����
    '=========================================================================
        
    '[�ʼ�] ������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    Taxinvoice.invoicerCorpNum = txtCorpNum.Text
    
    '[�ʼ�] ������ ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    Taxinvoice.invoicerTaxRegID = ""
    
    '[�ʼ�] ������ ��ȣ
    Taxinvoice.invoicerCorpName = "������ ��ȣ"
    
    '[�ʼ�] ������ ������ȣ, �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    Taxinvoice.invoicerMgtKey = txtMgtKey.Text
    
    '[�ʼ�] ������ ��ǥ�� ����
    Taxinvoice.invoicerCEOName = "������ ��ǥ�� ����"
    
    '������ �ּ�
    Taxinvoice.invoicerAddr = "������ �ּ�"
    
    '������ ����
    Taxinvoice.invoicerBizType = "������ ����,����2"
    
    '������ ����
    Taxinvoice.invoicerBizClass = "������ ����"
    
    '������ ����ڸ�
    Taxinvoice.invoicerContactName = "������ ����ڸ�"
    
    '������ ����� �����ּ�
    Taxinvoice.invoicerEmail = "test@test.com"
    
    '������ ����� ����ó
    Taxinvoice.invoicerTEL = "070-7070-0707"
    
    '������ ����� �޴�����ȣ
    Taxinvoice.invoicerHP = "010-000-2222"
    
    ' ����� �˸����� ���ۿ��� (�����࿡���� ��밡��)
    ' - ���޹޴��� ��)����� �޴�����ȣ(invoiceeHP1)�� ����
    ' - ���۽� ����Ʈ�� �����Ǹ� ���۽����ϴ� ��� ����Ʈ ȯ��ó��
    Taxinvoice.invoicerSMSSendYN = False
    
    
    '=========================================================================
    '                            ���޹޴��� ����
    '=========================================================================
        
    '[�ʼ�] ���޹޴��� ����, [�����, ����, �ܱ���] �� ����
    Taxinvoice.invoiceeType = "�����"
    
    '[�ʼ�] ���޹޴��� ����ڹ�ȣ, '-' ���� 10�ڸ�
    Taxinvoice.invoiceeCorpNum = "8888888888"
    
    '[�ʼ�] ���޹޴��� ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    Taxinvoice.invoiceeTaxRegID = ""
    
    '[�ʼ�] �����ڹ޴��� ��ȣ
    Taxinvoice.invoiceeCorpName = "���޹޴��� ��ȣ"
    
    '[������� �ʼ�] ���޹޴��� ������ȣ,  �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    Taxinvoice.invoiceeMgtKey = ""
    
    '[�ʼ�] ���޹޴��� ��ǥ�� ����
    Taxinvoice.invoiceeCEOName = "���޹޴��� ��ǥ�� ����"
    
    '���޹޴��� �ּ�
    Taxinvoice.invoiceeAddr = "���޹޴��� �ּ�"
    
    '���޹޴��� ����
    Taxinvoice.invoiceeBizClass = "���޹޴��� ����"
    
    '���޹޴��� ����
    Taxinvoice.invoiceeBizType = "���޹޴��� ����"
    
    '���޹޴��� ����ڸ�
    Taxinvoice.invoiceeContactName1 = "���޹޴��� ����ڸ�"
    
    '���޹޴��� ����� �����ּ�
    '�˺� ����ȯ�濡�� �׽�Ʈ�ϴ� ��쿡�� �ȳ� ������ ���۵ǹǷ�,
    '���� �ŷ�ó�� �����ּҰ� ������� �ʵ��� ����
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    '���޹޴��� ����� ����ó
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    
    '���޹޴��� ����� �޴�����ȣ
    Taxinvoice.invoiceeHP1 = "010-111-222"
    
    
    '=========================================================================
    '                            ���ݰ�꼭 ����
    '=========================================================================
    
    '[�ʼ�] ���ް��� �հ�
    Taxinvoice.supplyCostTotal = "200000"
    
    '[�ʼ�] ���� �հ�
    Taxinvoice.taxTotal = "20000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + �����հ�
    Taxinvoice.totalAmount = "220000"
    
    '���� �� '�Ϸù�ȣ' �׸�
    Taxinvoice.serialNum = "123"
    
    '���� �� '��' �׸�, �ִ밪 32767
    '�̱���� Taxinvoice.kwon = ""
    Taxinvoice.kwon = "1"
    
    '���� �� 'ȣ' �׸�, �ִ밪 32767
    '�̱���� Taxinvoice.kwon = ""
    Taxinvoice.ho = "1"
    
    '���� �� '����' �׸�
    Taxinvoice.cash = ""
    
    '���� �� '��ǥ' �׸�
    Taxinvoice.chkBill = ""
    
    '���� �� '����' �׸�
    Taxinvoice.note = ""
    
    '���� �� '�ܻ�̼���' �׸�
    Taxinvoice.credit = ""
    
    '���� �� '���'�׸�
    Taxinvoice.remark1 = "���1"
    Taxinvoice.remark2 = "���2"
    Taxinvoice.remark3 = "���3"
    
    '����ڵ���� �̹��� ÷�ο���
    Taxinvoice.businessLicenseYN = False
    
    '����纻 �̹��� ÷�ο���
    Taxinvoice.bankBookYN = False
    

    '=========================================================================
    '         �������ݰ�꼭 ���� (�������ݰ�꼭 �ۼ��ÿ��� ����
    ' - �������ݰ�꼭 ���� ������ �����Ŵ��� �Ǵ� ���߰��̵� ��ũ ����
    ' - [����] �������ݰ�꼭 �ۼ���� �ȳ� - https://developers.popbill.com/guide/taxinvoice/vb/introduction/modified-taxinvoice
    '========================================================================='
    
    ' ���������ڵ�, ���������� ���� 1~6�� ���ñ���
    Taxinvoice.modifyCode = ""
    
    ' �������ݰ�꼭 ����û���ι�ȣ ����
    Taxinvoice.orgNTSConfirmNum = ""
        
    
    '=========================================================================
    '             ���׸�(ǰ��) ���� > �迭�� 99������ ���� ����
    '=========================================================================
    
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20220101"   '�ŷ�����  yyyyMMdd
    newDetail.itemName = "ǰ��"         'ǰ���
    newDetail.spec = "�԰�"             '�԰�
    newDetail.qty = "1"                 '����
    newDetail.unitCost = "100000"       '�ܰ�
    newDetail.supplyCost = "100000"     '���ް���
    newDetail.tax = "10000"             '����
    newDetail.remark = "���"           '���
    
    Taxinvoice.detailList.Add newDetail
    
    Dim newDetail2 As New PBTIDetail
    newDetail2.serialNum = 2             '�Ϸù�ȣ 1���� ���� ����
    newDetail2.purchaseDT = "20220101"   '�ŷ�����  yyyyMMdd
    newDetail2.itemName = "ǰ��2"        'ǰ���
    newDetail2.spec = "�԰�"             '�԰�
    newDetail2.qty = "1"                 '����
    newDetail2.unitCost = "100000"       '�ܰ�
    newDetail2.supplyCost = "100000"     '���ް���
    newDetail2.tax = "10000"             '����
    newDetail2.remark = "���"           '���
    
    Taxinvoice.detailList.Add newDetail2
    
    
    '=========================================================================
    '             �߰������ ���� > �迭�� 5������ ���� ����
    ' - ���ݰ�꼭 ����ȳ� ������ ���Ź��� ���޹޴��� ����ڰ� �ټ��� ���
    ' ����� ������ �߰��Ͽ� ����ȳ������� �ټ����� ������ �� �ֽ��ϴ�.
    '=========================================================================
    
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                 '�Ϸù�ȣ, 1���� ��������
    newContact.ContactName = "����� ����"   '����ڸ�
    newContact.email = "test2@test.com"      '����� �����ּ�
    Taxinvoice.addContactList.Add newContact
    
    Dim newContact2 As New PBTIContact
    newContact2.serialNum = 2                '�Ϸù�ȣ, 1���� ��������
    newContact2.ContactName = "����� ����"  '����ڸ�
    newContact2.email = "test2@test.com"     '����� �����ּ�
    Taxinvoice.addContactList.Add newContact2
        
    
    '�ŷ����� �����ۼ� ����
    Taxinvoice.writeSpecification = False
    
    '�ŷ����� �����ۼ��� �ŷ����� ������ȣ, �̱���� ���ݰ�꼭 ������ȣ�� �ڵ��ۼ�
    Taxinvoice.dealInvoiceMgtKey = ""
    
    '�������� ��������(forceIssue)
    '���ึ������ ���� ���ݰ�꼭�� �����ϴ� ���, ���꼼�� �ΰ��� �� �ֽ��ϴ�.
    '���꼼�� �ΰ��Ǵ��� ������ �ؾ��ϴ� ��쿡�� forceIssue�� ����
    'true�� �����Ͽ� ����(Issue API)�� ȣ���Ͻø� �˴ϴ�.
    Taxinvoice.forceIssue = False
    
    '�޸�
    Taxinvoice.memo = ""
    
    '����ȳ� ��������, ����ó���� �⺻�������� ����
    Taxinvoice.emailSubject = ""
    
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.RegistIssue(txtCorpNum.Text, Taxinvoice)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message + vbCrLf + "����û���ι�ȣ : " + Response.ntsConfirmNum)
End Sub

'=========================================================================
' �ִ� 100���� ���ݰ�꼭 ������ �ѹ��� ��û���� �����մϴ�.
' - ���ݰ�꼭 ������ ���ؼ� �������� �������� �˺� ���������� ������� �Ǿ�� �մϴ�.
'    �� ����Ź������ ���, ��Ź���� ������ ����� �ʿ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#BulkSubmit
'=========================================================================
Private Sub btnBulkSubmit_Click()
    Dim Response As PBBulkResponse
    Dim taxinvoiceList As New Collection
    
    Dim i As Integer
    For i = 0 To 50
        Dim Taxinvoice
        Set Taxinvoice = New PBTaxinvoice
        '[�ʼ�] �ۼ�����, ǥ������ (yyyyMMdd)
        Taxinvoice.writeDate = "20220101"
        
        '[�ʼ�] ��������, [������, ������, ����Ź] �� ����
        Taxinvoice.issueType = "������"
        
        '[�ʼ�] {������, ������} �� ����, '������'�� ������ ���μ��������� �̿밡��
        '- ������(������ ����), ������(���޹޴��� ����)
        Taxinvoice.chargeDirection = "������"
        
        '[�ʼ�] ����/û��, [����, û��, ����] �� ����
        Taxinvoice.purposeType = "����"
            
        '[�ʼ�] ��������, [����, ����, �鼼] �� ����
        Taxinvoice.taxType = "����"
        
        
        '=========================================================================
        '                              ������ ����
        '=========================================================================
            
        '[�ʼ�] ������ ����ڹ�ȣ, '-' ���� 10�ڸ�
        Taxinvoice.invoicerCorpNum = txtCorpNum.Text
        
        '[�ʼ�] ������ ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
        Taxinvoice.invoicerTaxRegID = ""
        
        '[�ʼ�] ������ ��ȣ
        Taxinvoice.invoicerCorpName = "������ ��ȣ"
        
        '[�ʼ�] ������ ������ȣ, �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
        Taxinvoice.invoicerMgtKey = txtSubmitID.Text + CStr(i)
        
        '[�ʼ�] ������ ��ǥ�� ����
        Taxinvoice.invoicerCEOName = "������ ��ǥ�� ����"
        
        '������ �ּ�
        Taxinvoice.invoicerAddr = "������ �ּ�"
        
        '������ ����
        Taxinvoice.invoicerBizType = "������ ����,����2"
        
        '������ ����
        Taxinvoice.invoicerBizClass = "������ ����"
        
        '������ ����ڸ�
        Taxinvoice.invoicerContactName = "������ ����ڸ�"
        
        '������ ����� �����ּ�
        Taxinvoice.invoicerEmail = "test@test.com"
        
        '������ ����� ����ó
        Taxinvoice.invoicerTEL = "070-7070-0707"
        
        '������ ����� �޴�����ȣ
        Taxinvoice.invoicerHP = "010-000-2222"
        
        ' ����� �˸����� ���ۿ��� (�����࿡���� ��밡��)
        ' - ���޹޴��� ��)����� �޴�����ȣ(invoiceeHP1)�� ����
        ' - ���۽� ����Ʈ�� �����Ǹ� ���۽����ϴ� ��� ����Ʈ ȯ��ó��
        Taxinvoice.invoicerSMSSendYN = False
        
        
        '=========================================================================
        '                            ���޹޴��� ����
        '=========================================================================
            
        '[�ʼ�] ���޹޴��� ����, [�����, ����, �ܱ���] �� ����
        Taxinvoice.invoiceeType = "�����"
        
        '[�ʼ�] ���޹޴��� ����ڹ�ȣ, '-' ���� 10�ڸ�
        Taxinvoice.invoiceeCorpNum = "8888888888"
        
        '[�ʼ�] ���޹޴��� ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
        Taxinvoice.invoiceeTaxRegID = ""
        
        '[�ʼ�] �����ڹ޴��� ��ȣ
        Taxinvoice.invoiceeCorpName = "���޹޴��� ��ȣ"
        
        '[������� �ʼ�] ���޹޴��� ������ȣ,  �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
        Taxinvoice.invoiceeMgtKey = ""
        
        '[�ʼ�] ���޹޴��� ��ǥ�� ����
        Taxinvoice.invoiceeCEOName = "���޹޴��� ��ǥ�� ����"
        
        '���޹޴��� �ּ�
        Taxinvoice.invoiceeAddr = "���޹޴��� �ּ�"
        
        '���޹޴��� ����
        Taxinvoice.invoiceeBizClass = "���޹޴��� ����"
        
        '���޹޴��� ����
        Taxinvoice.invoiceeBizType = "���޹޴��� ����"
        
        '���޹޴��� ����ڸ�
        Taxinvoice.invoiceeContactName1 = "���޹޴��� ����ڸ�"
        
        '���޹޴��� ����� �����ּ�
        '�˺� ����ȯ�濡�� �׽�Ʈ�ϴ� ��쿡�� �ȳ� ������ ���۵ǹǷ�,
        '���� �ŷ�ó�� �����ּҰ� ������� �ʵ��� ����
        Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
        
        '���޹޴��� ����� ����ó
        Taxinvoice.invoiceeTEL1 = "070-1234-1234"
        
        '���޹޴��� ����� �޴�����ȣ
        Taxinvoice.invoiceeHP1 = "010-111-222"
        
        
        '=========================================================================
        '                            ���ݰ�꼭 ����
        '=========================================================================
        
        '[�ʼ�] ���ް��� �հ�
        Taxinvoice.supplyCostTotal = "200000"
        
        '[�ʼ�] ���� �հ�
        Taxinvoice.taxTotal = "20000"
        
        '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + �����հ�
        Taxinvoice.totalAmount = "220000"
        
        '���� �� '�Ϸù�ȣ' �׸�
        Taxinvoice.serialNum = "123"
        
        '���� �� '��' �׸�, �ִ밪 32767
        '�̱���� Taxinvoice.kwon = ""
        Taxinvoice.kwon = "1"
        
        '���� �� 'ȣ' �׸�, �ִ밪 32767
        '�̱���� Taxinvoice.kwon = ""
        Taxinvoice.ho = "1"
        
        '���� �� '����' �׸�
        Taxinvoice.cash = ""
        
        '���� �� '��ǥ' �׸�
        Taxinvoice.chkBill = ""
        
        '���� �� '����' �׸�
        Taxinvoice.note = ""
        
        '���� �� '�ܻ�̼���' �׸�
        Taxinvoice.credit = ""
        
        '���� �� '���'�׸�
        Taxinvoice.remark1 = "���1"
        Taxinvoice.remark2 = "���2"
        Taxinvoice.remark3 = "���3"
        
        '����ڵ���� �̹��� ÷�ο���
        Taxinvoice.businessLicenseYN = False
        
        '����纻 �̹��� ÷�ο���
        Taxinvoice.bankBookYN = False
        
    
        '=========================================================================
        '         �������ݰ�꼭 ���� (�������ݰ�꼭 �ۼ��ÿ��� ����
        ' - �������ݰ�꼭 ���� ������ �����Ŵ��� �Ǵ� ���߰��̵� ��ũ ����
        ' - [����] �������ݰ�꼭 �ۼ���� �ȳ� - https://developers.popbill.com/guide/taxinvoice/vb/introduction/modified-taxinvoice
        '========================================================================='
        
        ' ���������ڵ�, ���������� ���� 1~6�� ���ñ���
        Taxinvoice.modifyCode = ""
        
        ' �������ݰ�꼭 ����û���ι�ȣ ����
        Taxinvoice.orgNTSConfirmNum = ""
            
        
        '=========================================================================
        '             ���׸�(ǰ��) ���� > �迭�� 99������ ���� ����
        '=========================================================================
        
        Set Taxinvoice.detailList = New Collection
        
        Dim newDetail As New PBTIDetail
        
        newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
        newDetail.purchaseDT = "20220101"   '�ŷ�����  yyyyMMdd
        newDetail.itemName = "ǰ��"         'ǰ���
        newDetail.spec = "�԰�"             '�԰�
        newDetail.qty = "1"                 '����
        newDetail.unitCost = "100000"       '�ܰ�
        newDetail.supplyCost = "100000"     '���ް���
        newDetail.tax = "10000"             '����
        newDetail.remark = "���"           '���
        
        Taxinvoice.detailList.Add newDetail
        
        Dim newDetail2 As New PBTIDetail
        newDetail2.serialNum = 2             '�Ϸù�ȣ 1���� ���� ����
        newDetail2.purchaseDT = "20220101"   '�ŷ�����  yyyyMMdd
        newDetail2.itemName = "ǰ��2"        'ǰ���
        newDetail2.spec = "�԰�"             '�԰�
        newDetail2.qty = "1"                 '����
        newDetail2.unitCost = "100000"       '�ܰ�
        newDetail2.supplyCost = "100000"     '���ް���
        newDetail2.tax = "10000"             '����
        newDetail2.remark = "���"           '���
        
        Taxinvoice.detailList.Add newDetail2
        
        
        '=========================================================================
        '             �߰������ ���� > �迭�� 5������ ���� ����
        ' - ���ݰ�꼭 ����ȳ� ������ ���Ź��� ���޹޴��� ����ڰ� �ټ��� ���
        ' ����� ������ �߰��Ͽ� ����ȳ������� �ټ����� ������ �� �ֽ��ϴ�.
        '=========================================================================
        
        Set Taxinvoice.addContactList = New Collection
        Dim newContact As New PBTIContact
        newContact.serialNum = 1                 '�Ϸù�ȣ, 1���� ��������
        newContact.ContactName = "����� ����"   '����ڸ�
        newContact.email = "test2@test.com"      '����� �����ּ�
        Taxinvoice.addContactList.Add newContact
        
        Dim newContact2 As New PBTIContact
        newContact2.serialNum = 2                '�Ϸù�ȣ, 1���� ��������
        newContact2.ContactName = "����� ����"  '����ڸ�
        newContact2.email = "test2@test.com"     '����� �����ּ�
        Taxinvoice.addContactList.Add newContact2
            
        taxinvoiceList.Add Taxinvoice
    Next

    Set Response = TaxinvoiceService.BulkSubmit(txtCorpNum.Text, txtSubmitID.Text, taxinvoiceList, False)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message + vbCrLf + "�������̵� : " + Response.receiptID)
End Sub

'=========================================================================
' ������ ������ SubmitID�� ����Ͽ� ���ݰ�꼭 ��������� Ȯ���մϴ�.
' - ���� ���ݰ�꼭 ó�����´� ��������(txState)�� �Ϸ�(2) �� ��ȯ�˴ϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#GetBulkResult
'=========================================================================
Private Sub btnGetBulkResult_Click()
    Dim Response As PBBulkTaxinvoiceResult
    Dim tmp As String
    
    Set Response = TaxinvoiceService.GetBulkResult(txtCorpNum.Text, txtSubmitID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "code (�����ڵ�) : " + CStr(Response.code) + vbCrLf
    tmp = tmp + "message (����޽���) : " + Response.message + vbCrLf
    tmp = tmp + "submitID (������̵�) : " + Response.submitID + vbCrLf
    tmp = tmp + "submitCount (���ݰ�꼭 ���� �Ǽ�) : " + CStr(Response.submitCount) + vbCrLf
    tmp = tmp + "successCount (���ݰ�꼭 ���� ���� �Ǽ�) : " + CStr(Response.successCount) + vbCrLf
    tmp = tmp + "failCount (���ݰ�꼭 ���� ���� �Ǽ�) : " + CStr(Response.failCount) + vbCrLf
    tmp = tmp + "txState (���������ڵ�) : " + CStr(Response.txState) + vbCrLf
    tmp = tmp + "txResultCode (���� ����ڵ�) : " + CStr(Response.txResultCode) + vbCrLf
    tmp = tmp + "txStartDT (����ó�� �����Ͻ�) : " + Response.txStartDT + vbCrLf
    tmp = tmp + "txEndDT (����ó�� �Ϸ��Ͻ�) : " + Response.txEndDT + vbCrLf
    tmp = tmp + "receiptDT (���� �����Ͻ�) : " + Response.receiptDT + vbCrLf
    tmp = tmp + "receiptID (�������̵�) : " + Response.receiptDT + vbCrLf
  
    
    tmp = tmp + "invoicerMgtKey(������ ������ȣ) |  code (�ڵ�) | message (�޽���) |  ntsconfirmNum (����û���ι�ȣ) |  issueDT (�����Ͻ�) " + vbCrLf + vbCrLf
            
    Dim issueResult As PBBulkTaxinvoiceIssueResult
    
    If Response.issueResult Is Nothing = False Then
        For Each issueResult In Response.issueResult
            tmp = tmp + issueResult.invoicerMgtKey + " | "
            tmp = tmp + CStr(issueResult.code) + " | "
            tmp = tmp + issueResult.message + " | "
            tmp = tmp + issueResult.ntsConfirmNum + " | "
            tmp = tmp + issueResult.issueDT + vbCrLf
        Next
    End If
    
    MsgBox tmp
End Sub

'=========================================================================
' ����û ���� ���� "����Ϸ�" ������ ���ݰ�꼭�� "�������"�ϰ� ����û ���� ��󿡼� �����մϴ�.
' - ����(Delete API) �Լ��� ȣ���Ͽ� "�������" ������ ���ڼ��ݰ�꼭�� �����ϸ�, ������ȣ ������ �����մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#CancelIssue
'=========================================================================
Private Sub btnCancelIssue_sub_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim memo As String
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    '�޸�
    memo = "���� ��� �޸�"
    
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���� ������ ������ ���ݰ�꼭�� �����մϴ�.
' - ���� ������ ����: "�ӽ�����", "�������", "������ź�", "���������", "���۽���"
' - ����ó���� ���ݰ�꼭�� ������ȣ�� ������ �����մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#Delete
'=========================================================================
Private Sub btnDelete_sub_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
        
    Set Response = TaxinvoiceService.Delete(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �ۼ��� ���ݰ�꼭 �����͸� �˺��� �����մϴ�.
' - "�ӽ�����" ������ ���ݰ�꼭�� ����(Issue) �Լ��� ȣ���Ͽ� "����Ϸ�" ó���� ��쿡�� ����û���� ���۵˴ϴ�.
' - ������ �� �ӽ�����(Register)�� ����(Issue)�� �ѹ��� ȣ��� ó���ϴ� ��ù���(RegistIssue API) ���μ��� ������ �����մϴ�.
' - ������ �� �ӽ�����(Register)�� �������û(Request)�� �ѹ��� ȣ��� ó���ϴ� ��ÿ�û(RegistRequest API) ���μ��� ������ �����մϴ�.
' - ���ݰ�꼭 ����÷�� ����� �����ϴ� ���, �ӽ�����(Register API) -> ����÷��(AttachFile API) -> ����(Issue API) �Լ��� ���ʷ� ȣ���մϴ�.
' - ������ ���ݰ�꼭�� �����ϴ� ���, ��ü 'Taxinvoice'�� ���� 'chargeDirection' ���� ���� ���� ��ü�� ������ �� �ֽ��ϴ�.
'   �� ������ : ������ ���� , ������ : ���޹޴��� ����
' - �ӽ������ ���ݰ�꼭�� �˺� ����Ʈ '�ӽù�����'���� Ȯ�� �����մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#Register
'=========================================================================

Private Sub btnRegister_Click()
    Dim writeSpecification As Boolean
    Dim Taxinvoice As New PBTaxinvoice
    
    '[�ʼ�] �ۼ�����, ǥ������ (yyyyMMdd)
    Taxinvoice.writeDate = "20220101"
    
    '[�ʼ�] ��������, [������, ������, ����Ź] �� ����
    Taxinvoice.issueType = "������"
    
    '[�ʼ�] {������, ������} �� ����, '������'�� ������ ���μ��������� �̿밡��
    '- ������(������ ����), ������(���޹޴��� ����)
    Taxinvoice.chargeDirection = "������"
    
    '[�ʼ�] ����/û��, [����, û��, ����] �� ����
    Taxinvoice.purposeType = "����"
        
    '[�ʼ�] ��������, [����, ����, �鼼] �� ����
    Taxinvoice.taxType = "����"
    
    
    '=========================================================================
    '                              ������ ����
    '=========================================================================
        
    '[�ʼ�] ������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    Taxinvoice.invoicerCorpNum = txtCorpNum.Text
    
    '[�ʼ�] ������ ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    Taxinvoice.invoicerTaxRegID = ""
    
    '[�ʼ�] ������ ��ȣ
    Taxinvoice.invoicerCorpName = "������ ��ȣ"
    
    '[�ʼ�] ������ ������ȣ, �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    Taxinvoice.invoicerMgtKey = txtMgtKey.Text
    
    '[�ʼ�] ������ ��ǥ�� ����
    Taxinvoice.invoicerCEOName = "������ ��ǥ�� ����"
    
    '������ �ּ�
    Taxinvoice.invoicerAddr = "������ �ּ�"
    
    '������ ����
    Taxinvoice.invoicerBizType = "������ ����,����2"
    
    '������ ����
    Taxinvoice.invoicerBizClass = "������ ����"
    
    '������ ����ڸ�
    Taxinvoice.invoicerContactName = "������ ����ڸ�"
    
    '������ ����� �����ּ�
    Taxinvoice.invoicerEmail = "test@test.com"
    
    '������ ����� ����ó
    Taxinvoice.invoicerTEL = "070-4304-2991"
    
    '������ ����� �޴�����ȣ
    Taxinvoice.invoicerHP = "010-000-2222"
    
    ' ����� �˸����� ���ۿ��� (�����࿡���� ��밡��)
    ' - ���޹޴��� ��)����� �޴�����ȣ(invoiceeHP1)�� ����
    ' - ���۽� ����Ʈ�� �����Ǹ� ���۽����ϴ� ��� ����Ʈ ȯ��ó��
    Taxinvoice.invoicerSMSSendYN = False
    
    
    '=========================================================================
    '                            ���޹޴��� ����
    '=========================================================================
        
    '[�ʼ�] ���޹޴��� ����, [�����, ����, �ܱ���] �� ����
    Taxinvoice.invoiceeType = "�����"
    
    '[�ʼ�] ���޹޴��� ����ڹ�ȣ, '-' ���� 10�ڸ�
    Taxinvoice.invoiceeCorpNum = "8888888888"
    
    '[�ʼ�] ���޹޴��� ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    Taxinvoice.invoiceeTaxRegID = ""
    
    '[�ʼ�] �����ڹ޴��� ��ȣ
    Taxinvoice.invoiceeCorpName = "���޹޴��� ��ȣ"
    
    '[������� �ʼ�] ���޹޴��� ������ȣ, �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    Taxinvoice.invoiceeMgtKey = ""
    
    '[�ʼ�] ���޹޴��� ��ǥ�� ����
    Taxinvoice.invoiceeCEOName = "���޹޴��� ��ǥ�� ����"
    
    '���޹޴��� �ּ�
    Taxinvoice.invoiceeAddr = "���޹޴��� �ּ�"
    
    '���޹޴��� ����
    Taxinvoice.invoiceeBizClass = "���޹޴��� ����"
    
    '���޹޴��� ����
    Taxinvoice.invoiceeBizType = "���޹޴��� ����"
    
    '���޹޴��� ����ڸ�
    Taxinvoice.invoiceeContactName1 = "���޹޴��� ����ڸ�"
    
    '���޹޴��� ����� �����ּ�
    '�˺� ����ȯ�濡�� �׽�Ʈ�ϴ� ��쿡�� �ȳ� ������ ���۵ǹǷ�,
    '���� �ŷ�ó�� �����ּҰ� ������� �ʵ��� ����
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    '���޹޴��� ����� ����ó
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    
    '���޹޴��� ����� �޴�����ȣ
    Taxinvoice.invoiceeHP1 = "010-111-222"
            
    
    '=========================================================================
    '                            ���ݰ�꼭 ����
    '=========================================================================
    
    '[�ʼ�] ���ް��� �հ�
    Taxinvoice.supplyCostTotal = "200000"
    
    '[�ʼ�] ���� �հ�
    Taxinvoice.taxTotal = "20000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + �����հ�
    Taxinvoice.totalAmount = "220000"
    
    '���� �� '�Ϸù�ȣ' �׸�
    Taxinvoice.serialNum = "123"
    
    '���� �� '��' �׸�, �ִ밪 32767
    ' �̱���� Taxinvoice.kwon = ""
    Taxinvoice.kwon = "1"
    
    '���� �� 'ȣ' �׸�, �ִ밪 32767
    ' �̱���� Taxinvoice.kwon = ""
    Taxinvoice.ho = "1"
    
    '���� �� '����' �׸�
    Taxinvoice.cash = ""
    
    '���� �� '��ǥ' �׸�
    Taxinvoice.chkBill = ""
    
    '���� �� '����' �׸�
    Taxinvoice.note = ""
    
    '���� �� '�ܻ�̼���' �׸�
    Taxinvoice.credit = ""
    
    '���� �� '���'�׸�
    Taxinvoice.remark1 = "���1"
    Taxinvoice.remark2 = "���2"
    Taxinvoice.remark3 = "���3"
    
    '����ڵ���� �̹��� ÷�ο���
    Taxinvoice.businessLicenseYN = False
    
    '����纻 �̹��� ÷�ο���
    Taxinvoice.bankBookYN = False
    

    '=========================================================================
    '         �������ݰ�꼭 ���� (�������ݰ�꼭 �ۼ��ÿ��� ����)
    ' - �������ݰ�꼭 ���� ������ �����Ŵ��� �Ǵ� ���߰��̵� ��ũ ����
    ' - [����] �������ݰ�꼭 �ۼ���� �ȳ� - https://developers.popbill.com/guide/taxinvoice/vb/introduction/modified-taxinvoice
    '========================================================================='
    
    ' ���������ڵ�, ���������� ���� 1~6�� ���ñ���
    Taxinvoice.modifyCode = ""
    
    ' �������ݰ�꼭 ����û���ι�ȣ ����
    Taxinvoice.orgNTSConfirmNum = ""
        
    
    '=========================================================================
    '             ���׸�(ǰ��) ���� > �迭�� 99������ ���� ����
    '=========================================================================
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20220101"   '�ŷ�����  yyyyMMdd
    newDetail.itemName = "ǰ��"         'ǰ���
    newDetail.spec = "�԰�"             '�԰�
    newDetail.qty = "1"                 '����
    newDetail.unitCost = "100000"       '�ܰ�
    newDetail.supplyCost = "100000"     '���ް���
    newDetail.tax = "10000"             '����
    newDetail.remark = "���"           '���
    
    Taxinvoice.detailList.Add newDetail
    
    Dim newDetail2 As New PBTIDetail
    newDetail2.serialNum = 2             '�Ϸù�ȣ 1���� ���� ����
    newDetail2.purchaseDT = "20220101"   '�ŷ�����  yyyyMMdd
    newDetail2.itemName = "ǰ��2"        'ǰ���
    newDetail2.spec = "�԰�"             '�԰�
    newDetail2.qty = "1"                 '����
    newDetail2.unitCost = "100000"       '�ܰ�
    newDetail2.supplyCost = "100000"     '���ް���
    newDetail2.tax = "10000"             '����
    newDetail2.remark = "���"           '���
    
    Taxinvoice.detailList.Add newDetail2
    
    
    '=========================================================================
    '             �߰������ ���� > �迭�� 5������ ���� ����
    ' - ���ݰ�꼭 ����ȳ� ������ ���Ź��� ���޹޴��� ����ڰ� �ټ��� ���
    ' ����� ������ �߰��Ͽ� ����ȳ������� �ټ����� ������ �� �ֽ��ϴ�.
    '=========================================================================
    Set Taxinvoice.addContactList = New Collection
    
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                 '�Ϸù�ȣ, 1���� ��������
    newContact.ContactName = "����� ����"   '����ڸ�
    newContact.email = "test2@test.com"      '����� �����ּ�
    Taxinvoice.addContactList.Add newContact
    
    Dim newContact2 As New PBTIContact
    newContact2.serialNum = 2                '�Ϸù�ȣ, 1���� ��������
    newContact2.ContactName = "����� ����"  '����ڸ�
    newContact2.email = "test2@test.com"     '����� �����ּ�
    Taxinvoice.addContactList.Add newContact2
    
    '�ŷ����� �����ۼ� ����
    writeSpecification = False
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.Register(txtCorpNum.Text, Taxinvoice, writeSpecification)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' "�ӽ�����" ������ ���ݰ�꼭�� �����մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#Update
'=========================================================================
Private Sub btnUpdate_Click()
    Dim KeyType As MgtKeyType
    Dim writeSpecification As Boolean
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    Dim Taxinvoice As New PBTaxinvoice
    
    '[�ʼ�] �ۼ�����, ǥ������ (yyyyMMdd)
    Taxinvoice.writeDate = "20220101"
    
    '[�ʼ�] ��������, [������, ������, ����Ź] �� ����
    Taxinvoice.issueType = "������"
    
    '[�ʼ�] {������, ������} �� ����, '������'�� ������ ���μ��������� �̿밡��
    '- ������(������ ����), ������(���޹޴��� ����)
    Taxinvoice.chargeDirection = "������"
    
    '[�ʼ�] ����/û��, [����, û��, ����] �� ����
    Taxinvoice.purposeType = "����"
        
    '[�ʼ�] ��������, [����, ����, �鼼] �� ����
    Taxinvoice.taxType = "����"
    
    
    '=========================================================================
    '                              ������ ����
    '=========================================================================
        
    '[�ʼ�] ������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    Taxinvoice.invoicerCorpNum = txtCorpNum.Text
    
    '[�ʼ�] ������ ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    Taxinvoice.invoicerTaxRegID = ""
    
    '[�ʼ�] ������ ��ȣ
    Taxinvoice.invoicerCorpName = "������ ��ȣ_����"
    
    '[�ʼ�] ������ ������ȣ, �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    Taxinvoice.invoicerMgtKey = txtMgtKey.Text
    
    '[�ʼ�] ������ ��ǥ�� ����
    Taxinvoice.invoicerCEOName = "������ ��ǥ�� ����_����"
    
    '������ �ּ�
    Taxinvoice.invoicerAddr = "������ �ּ�"
    
    '������ ����
    Taxinvoice.invoicerBizType = "������ ����,����2"
    
    '������ ����
    Taxinvoice.invoicerBizClass = "������ ����"
    
    '������ ����ڸ�
    Taxinvoice.invoicerContactName = "������ ����ڸ�"
    
    '������ ����� �����ּ�
    Taxinvoice.invoicerEmail = "test@test.com"
    
    '������ ����� ����ó
    Taxinvoice.invoicerTEL = "070-7070-0707"
    
    '������ ����� �޴�����ȣ
    Taxinvoice.invoicerHP = "010-000-2222"
    
    ' ����� �˸����� ���ۿ��� (�����࿡���� ��밡��)
    ' - ���޹޴��� ��)����� �޴�����ȣ(invoiceeHP1)�� ����
    ' - ���۽� ����Ʈ�� �����Ǹ� ���۽����ϴ� ��� ����Ʈ ȯ��ó��
    Taxinvoice.invoicerSMSSendYN = False
    
    
    '=========================================================================
    '                            ���޹޴��� ����
    '=========================================================================
        
    '[�ʼ�] ���޹޴��� ����, [�����, ����, �ܱ���] �� ����
    Taxinvoice.invoiceeType = "�����"
    
    '[�ʼ�] ���޹޴��� ����ڹ�ȣ, '-' ���� 10�ڸ�
    Taxinvoice.invoiceeCorpNum = "8888888888"
    
    '[�ʼ�] ���޹޴��� ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    Taxinvoice.invoiceeTaxRegID = ""
    
    '[�ʼ�] �����ڹ޴��� ��ȣ
    Taxinvoice.invoiceeCorpName = "���޹޴��� ��ȣ"
    
    '[������� �ʼ�] ���޹޴��� ������ȣ, �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    Taxinvoice.invoiceeMgtKey = ""
    
    '[�ʼ�] ���޹޴��� ��ǥ�� ����
    Taxinvoice.invoiceeCEOName = "���޹޴��� ��ǥ�� ����"
    
    '���޹޴��� �ּ�
    Taxinvoice.invoiceeAddr = "���޹޴��� �ּ�"
    
    '���޹޴��� ����
    Taxinvoice.invoiceeBizClass = "���޹޴��� ����"
    
    '���޹޴��� ����
    Taxinvoice.invoiceeBizType = "���޹޴��� ����"
    
    '���޹޴��� ����ڸ�
    Taxinvoice.invoiceeContactName1 = "���޹޴��� ����ڸ�"
    
    '���޹޴��� ����� �����ּ�
    '�˺� ����ȯ�濡�� �׽�Ʈ�ϴ� ��쿡�� �ȳ� ������ ���۵ǹǷ�,
    '���� �ŷ�ó�� �����ּҰ� ������� �ʵ��� ����
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    '���޹޴��� ����� ����ó
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    
    '���޹޴��� ����� �޴�����ȣ
    Taxinvoice.invoiceeHP1 = "010-111-222"
    
    '=========================================================================
    '                            ���ݰ�꼭 ����
    '=========================================================================
    
    '[�ʼ�] ���ް��� �հ�
    Taxinvoice.supplyCostTotal = "200000"
    
    '[�ʼ�] ���� �հ�
    Taxinvoice.taxTotal = "20000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + �����հ�
    Taxinvoice.totalAmount = "220000"
    
    '���� �� '�Ϸù�ȣ' �׸�
    Taxinvoice.serialNum = "123"
    
    '���� �� '��' �׸�, �ִ밪 32767
    '�̱���� Taxinvoice.kwon = ""
    Taxinvoice.kwon = "1"
    
    '���� �� 'ȣ' �׸�, �ִ밪 32767
    '�̱���� Taxinvoice.kwon = ""
    Taxinvoice.ho = "1"
    
    '���� �� '����' �׸�
    Taxinvoice.cash = ""
    
    '���� �� '��ǥ' �׸�
    Taxinvoice.chkBill = ""
    
    '���� �� '����' �׸�
    Taxinvoice.note = ""
    
    '���� �� '�ܻ�̼���' �׸�
    Taxinvoice.credit = ""
    
    '���� �� '���'�׸�
    Taxinvoice.remark1 = "���1"
    Taxinvoice.remark2 = "���2"
    Taxinvoice.remark3 = "���3"
    
    '����ڵ���� �̹��� ÷�ο���
    Taxinvoice.businessLicenseYN = False
    
    '����纻 �̹��� ÷�ο���
    Taxinvoice.bankBookYN = False
    

    '=========================================================================
    '         �������ݰ�꼭 ���� (�������ݰ�꼭 �ۼ��ÿ��� ����
    ' - �������ݰ�꼭 ���� ������ �����Ŵ��� �Ǵ� ���߰��̵� ��ũ ����
    ' - [����] �������ݰ�꼭 �ۼ���� �ȳ� - https://developers.popbill.com/guide/taxinvoice/vb/introduction/modified-taxinvoice
    '========================================================================='
    
    ' ���������ڵ�, ���������� ���� 1~6�� ���ñ���
    Taxinvoice.modifyCode = ""
    
    ' �������ݰ�꼭 ����û���ι�ȣ ����
    Taxinvoice.orgNTSConfirmNum = ""
        
    
    '=========================================================================
    '             ���׸�(ǰ��) ���� > �迭�� 99������ ���� ����
    '=========================================================================
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20220101"   '�ŷ�����  yyyyMMdd
    newDetail.itemName = "ǰ��"         'ǰ���
    newDetail.spec = "�԰�"             '�԰�
    newDetail.qty = "1"                 '����
    newDetail.unitCost = "100000"       '�ܰ�
    newDetail.supplyCost = "100000"     '���ް���
    newDetail.tax = "10000"             '����
    newDetail.remark = "���"           '���
    
    Taxinvoice.detailList.Add newDetail
    
    Dim newDetail2 As New PBTIDetail
    newDetail2.serialNum = 2             '�Ϸù�ȣ 1���� ���� ����
    newDetail2.purchaseDT = "20220101"   '�ŷ�����  yyyyMMdd
    newDetail2.itemName = "ǰ��2"        'ǰ���
    newDetail2.spec = "�԰�"             '�԰�
    newDetail2.qty = "1"                 '����
    newDetail2.unitCost = "100000"       '�ܰ�
    newDetail2.supplyCost = "100000"     '���ް���
    newDetail2.tax = "10000"             '����
    newDetail2.remark = "���"           '���
    
    Taxinvoice.detailList.Add newDetail2
    
    
    '=========================================================================
    '             �߰������ ���� > �迭�� 5������ ���� ����
    ' - ���ݰ�꼭 ����ȳ� ������ ���Ź��� ���޹޴��� ����ڰ� �ټ��� ���
    ' ����� ������ �߰��Ͽ� ����ȳ������� �ټ����� ������ �� �ֽ��ϴ�.
    '=========================================================================
    Set Taxinvoice.addContactList = New Collection
    
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                 '�Ϸù�ȣ, 1���� ��������
    newContact.ContactName = "����� ����"   '����ڸ�
    newContact.email = "test2@test.com"      '����� �����ּ�
    Taxinvoice.addContactList.Add newContact
    
    Dim newContact2 As New PBTIContact
    newContact2.serialNum = 2                '�Ϸù�ȣ, 1���� ��������
    newContact2.ContactName = "����� ����"  '����ڸ�
    newContact2.email = "test2@test.com"     '����� �����ּ�
    
    Taxinvoice.addContactList.Add newContact2
    
    '�ŷ����� �����ۼ�����
    writeSpecification = False
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.Update(txtCorpNum.Text, KeyType, txtMgtKey.Text, Taxinvoice, writeSpecification)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'======================================================================================================================
' "�ӽ�����" �Ǵ� "(��)������" ������ ���ݰ�꼭�� ����(���ڼ���)�ϸ�, "����Ϸ�" ���·� ó���մϴ�.
' - ���ݰ�꼭 ����û ������å [https://developers.popbill.com/guide/taxinvoice/vb/introduction/policy-of-send-to-nts]
' - "����Ϸ�" �� ���ڼ��ݰ�꼭�� ����û ���� ������ �������(CancelIssue API) �Լ��� ����û �Ű� ��󿡼� ������ �� �ֽ��ϴ�.
' - ���ݰ�꼭 ������ ���ؼ� �������� �������� �˺� ���������� ������� �Ǿ�� �մϴ�.
'   �� ����Ź������ ���, ��Ź���� ������ ����� �ʿ��մϴ�.
' - ���ݰ�꼭 ���� �� ���޹޴��ڿ��� ���� ������ �߼۵˴ϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#Issue
'======================================================================================================================
Private Sub btnIssue_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim memo As String
    Dim emailSubject As String
    Dim forceIssue As Boolean
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    '�޸�
    memo = "�޸�"
    
    '���޹޴��ڿ��� ���۵Ǵ� ����ȳ����� ����, �̱���� �⺻������� ����
    emailSubject = ""
    
    '�������� ��������, �⺻�� - False
    '���ึ������ ���� ���ݰ�꼭�� �����ϴ� ���, ���꼼�� �ΰ��� �� �ֽ��ϴ�.
    '�������� ���ݰ�꼭�� �Ű��ؾ� �ϴ� ��� forceIssue ���� True�� �����Ͽ� ����(Issue API)�� ȣ���� �� �ֽ��ϴ�.
    forceIssue = False
        
    
    Set Response = TaxinvoiceService.Issue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, emailSubject, forceIssue)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message + vbCrLf + "����û���ι�ȣ : " + Response.ntsConfirmNum)
End Sub

'=========================================================================
' ����û ���� ���� "����Ϸ�" ������ ���ݰ�꼭�� "�������"�ϰ� ����û ���� ��󿡼� �����մϴ�.
' - ����(Delete API) �Լ��� ȣ���Ͽ� "�������" ������ ���ڼ��ݰ�꼭�� �����ϸ�, ������ȣ ������ �����մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#CancelIssue
'=========================================================================
Private Sub btnCancelIssue_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim memo As String
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    '�޸�
    memo = "������� �޸�"
    
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���� ������ ������ ���ݰ�꼭�� �����մϴ�.
' - ���� ������ ����: "�ӽ�����", "�������", "������ź�", "���������", "���۽���"
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#Delete
'=========================================================================
Private Sub btnDelete_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.Delete(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' "����Ϸ�" ������ ���ڼ��ݰ�꼭�� ����û�� ��� �����ϸ�, �Լ� ȣ�� �� �ִ� 30�� �̳��� ���� ó���� �Ϸ�˴ϴ�.
' - ����û ��������� ȣ������ ���� ���ݰ�꼭�� ������ ���� ���� ������ ���� 3�ÿ� �˺� �ý��ۿ��� �ϰ������� ����û���� �����մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#SendToNTS
'=========================================================================
Private Sub btnSendToNTS_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.sendToNTS(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================================================================================
' ���޹޴��ڰ� �ۼ��� ���ݰ�꼭 �����͸� �˺��� �����ϰ� �����ڿ��� �ۺ��Ͽ� ������ ��û�մϴ�.
' - ������ ���ݰ�꼭 ���μ����� �����ϱ� ���ؼ��� ������/���޹޴��ڰ� ��� �˺��� ȸ���̿��� �մϴ�.
' - ���� ��û�� ���ݰ�꼭�� "(��)������" �����̸�, �����ڰ� �˺� ����Ʈ �Ǵ� �Լ��� ȣ���Ͽ� ������ ��쿡�� ����û���� ���۵˴ϴ�.
' - �����ڴ� �˺� ����Ʈ�� "���� ���� �����"���� ������ ������ ������ ���ݰ�꼭�� Ȯ���� �� �ֽ��ϴ�.
' - �ӽ�����(Register API) �Լ��� ������ ��û(Request API) �Լ��� �� ���� ���μ����� ó���մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#RegistRequest
'=========================================================================================================================================
Private Sub btnRegistRequest_Click()
    Dim Taxinvoice As New PBTaxinvoice
    
    '[�ʼ�] �ۼ�����, ǥ������ (yyyyMMdd)
    Taxinvoice.writeDate = "20220101"
    
    '[�ʼ�] ��������, [������, ������, ����Ź] �� ����
    Taxinvoice.issueType = "������"
    
    '[�ʼ�] {������, ������} �� ����, '������'�� ������ ���μ��������� �̿밡��
    '- ������(������ ����), ������(���޹޴��� ����)
    Taxinvoice.chargeDirection = "������"
    
    '[�ʼ�] ����/û��, [����, û��, ����] �� ����
    Taxinvoice.purposeType = "����"
        
    '[�ʼ�] ��������, [����, ����, �鼼] �� ����
    Taxinvoice.taxType = "����"
    
    
    '=========================================================================
    '                              ������ ����
    '=========================================================================
        
    '[�ʼ�] ������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    Taxinvoice.invoicerCorpNum = "8888888888"
    
    '[�ʼ�] ������ ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    Taxinvoice.invoicerTaxRegID = ""
    
    '[�ʼ�] ������ ��ȣ
    Taxinvoice.invoicerCorpName = "������ ��ȣ"
    
    '[������� �ʼ�] ������ ������ȣ, �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    Taxinvoice.invoicerMgtKey = ""
    
    '[�ʼ�] ������ ��ǥ�� ����
    Taxinvoice.invoicerCEOName = "������ ��ǥ�� ����"
    
    '������ �ּ�
    Taxinvoice.invoicerAddr = "������ �ּ�"
    
    '������ ����
    Taxinvoice.invoicerBizType = "������ ����,����2"
    
    '������ ����
    Taxinvoice.invoicerBizClass = "������ ����"
    
    '������ ����ڸ�
    Taxinvoice.invoicerContactName = "������ ����ڸ�"
    
    '������ ����� �����ּ�
    Taxinvoice.invoicerEmail = "test@test.com"
    
    '������ ����� ����ó
    Taxinvoice.invoicerTEL = "070-7070-0707"
    
    '������ ����� �޴�����ȣ
    Taxinvoice.invoicerHP = "010-000-2222"
    
    
    '=========================================================================
    '                            ���޹޴��� ����
    '=========================================================================
        
    '[�ʼ�] ���޹޴��� ����, [�����, ����, �ܱ���] �� ����
    Taxinvoice.invoiceeType = "�����"
    
    '[�ʼ�] ���޹޴��� ����ڹ�ȣ, '-' ���� 10�ڸ�
    Taxinvoice.invoiceeCorpNum = txtCorpNum.Text
    
    '[�ʼ�] ���޹޴��� ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    Taxinvoice.invoiceeTaxRegID = ""
    
    '[�ʼ�] �����ڹ޴��� ��ȣ
    Taxinvoice.invoiceeCorpName = "���޹޴��� ��ȣ"
    
    '[������� �ʼ�] ���޹޴��� ������ȣ, �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    Taxinvoice.invoiceeMgtKey = txtMgtKey.Text
    
    '[�ʼ�] ���޹޴��� ��ǥ�� ����
    Taxinvoice.invoiceeCEOName = "���޹޴��� ��ǥ�� ����"
    
    '���޹޴��� �ּ�
    Taxinvoice.invoiceeAddr = "���޹޴��� �ּ�"
    
    '���޹޴��� ����
    Taxinvoice.invoiceeBizClass = "���޹޴��� ����"
    
    '���޹޴��� ����
    Taxinvoice.invoiceeBizType = "���޹޴��� ����"
    
    '���޹޴��� ����ڸ�
    Taxinvoice.invoiceeContactName1 = "���޹޴��� ����ڸ�"
    
    '���޹޴��� ����� �����ּ�
    '�˺� ����ȯ�濡�� �׽�Ʈ�ϴ� ��쿡�� �ȳ� ������ ���۵ǹǷ�,
    '���� �ŷ�ó�� �����ּҰ� ������� �ʵ��� ����
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    '���޹޴��� ����� ����ó
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    
    '���޹޴��� ����� �޴�����ȣ
    Taxinvoice.invoiceeHP1 = "010-111-222"
    
    ' ������ ��û�� �˸����� ���ۿ��� (�����࿡���� ��밡��)
    ' - ������ ����� �޴�����ȣ(invoicerHP)�� ����
    ' - ���۽� ����Ʈ�� �����Ǹ� ���۽����ϴ� ��� ����Ʈ ȯ��ó��
    Taxinvoice.invoiceeSMSSendYN = False
    
    '=========================================================================
    '                            ���ݰ�꼭 ����
    '=========================================================================
    
    '[�ʼ�] ���ް��� �հ�
    Taxinvoice.supplyCostTotal = "200000"
    
    '[�ʼ�] ���� �հ�
    Taxinvoice.taxTotal = "20000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + �����հ�
    Taxinvoice.totalAmount = "220000"
    
    '���� �� '�Ϸù�ȣ' �׸�
    Taxinvoice.serialNum = "123"
    
    '���� �� '��' �׸�, �ִ밪 32767
    '�̱���� Taxinvoice.kwon = ""
    Taxinvoice.kwon = "1"
    
    '���� �� 'ȣ' �׸�, �ִ밪 32767
    '�̱���� Taxinvoice.kwon = ""
    Taxinvoice.ho = "1"
    
    '���� �� '����' �׸�
    Taxinvoice.cash = ""
    
    '���� �� '��ǥ' �׸�
    Taxinvoice.chkBill = ""
    
    '���� �� '����' �׸�
    Taxinvoice.note = ""
    
    '���� �� '�ܻ�̼���' �׸�
    Taxinvoice.credit = ""
    
    '���� �� '���'�׸�
    Taxinvoice.remark1 = "���1"
    Taxinvoice.remark2 = "���2"
    Taxinvoice.remark3 = "���3"
    
    '����ڵ���� �̹��� ÷�ο���
    Taxinvoice.businessLicenseYN = False
    
    '����纻 �̹��� ÷�ο���
    Taxinvoice.bankBookYN = False
    

    '=========================================================================
    '         �������ݰ�꼭 ���� (�������ݰ�꼭 �ۼ��ÿ��� ����
    ' - �������ݰ�꼭 ���� ������ �����Ŵ��� �Ǵ� ���߰��̵� ��ũ ����
    ' - [����] �������ݰ�꼭 �ۼ���� �ȳ� - https://developers.popbill.com/guide/taxinvoice/vb/introduction/modified-taxinvoice
    '========================================================================='
    
    ' ���������ڵ�, ���������� ���� 1~6�� ���ñ���
    Taxinvoice.modifyCode = ""
    
    ' �������ݰ�꼭 ����û���ι�ȣ ����
    Taxinvoice.orgNTSConfirmNum = ""
        
    
    '=========================================================================
    '             ���׸�(ǰ��) ���� > �迭�� 99������ ���� ����
    '=========================================================================
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20220101"   '�ŷ�����  yyyyMMdd
    newDetail.itemName = "ǰ��"         'ǰ���
    newDetail.spec = "�԰�"             '�԰�
    newDetail.qty = "1"                 '����
    newDetail.unitCost = "100000"       '�ܰ�
    newDetail.supplyCost = "100000"     '���ް���
    newDetail.tax = "10000"             '����
    newDetail.remark = "���"           '���
    
    Taxinvoice.detailList.Add newDetail
    
    Dim newDetail2 As New PBTIDetail
    newDetail2.serialNum = 2             '�Ϸù�ȣ 1���� ���� ����
    newDetail2.purchaseDT = "20220101"   '�ŷ�����  yyyyMMdd
    newDetail2.itemName = "ǰ��2"        'ǰ���
    newDetail2.spec = "�԰�"             '�԰�
    newDetail2.qty = "1"                 '����
    newDetail2.unitCost = "100000"       '�ܰ�
    newDetail2.supplyCost = "100000"     '���ް���
    newDetail2.tax = "10000"             '����
    newDetail2.remark = "���"           '���
    
    Taxinvoice.detailList.Add newDetail2
        
    '�޸�
    Taxinvoice.memo = "��ÿ�û �޸�"
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.RegistRequest(txtCorpNum.Text, Taxinvoice)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'======================================================================================================================
' "�ӽ�����" �Ǵ� "(��)������" ������ ���ݰ�꼭�� ����(���ڼ���)�ϸ�, "����Ϸ�" ���·� ó���մϴ�.
' - ���ݰ�꼭 ����û ������å [https://developers.popbill.com/guide/taxinvoice/vb/introduction/policy-of-send-to-nts]
' - "����Ϸ�" �� ���ڼ��ݰ�꼭�� ����û ���� ������ �������(CancelIssue API) �Լ��� ����û �Ű� ��󿡼� ������ �� �ֽ��ϴ�.
' - ���ݰ�꼭 ������ ���ؼ� �������� �������� �˺� ���������� ������� �Ǿ�� �մϴ�.
'   �� ����Ź������ ���, ��Ź���� ������ ����� �ʿ��մϴ�.
' - ���ݰ�꼭 ���� �� ���޹޴��ڿ��� ���� ������ �߼۵˴ϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#Issue
'======================================================================================================================
Private Sub btnIssue_rev_sub_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim memo As String
    Dim emailSubject As String
    Dim forceIssue As Boolean
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    '�޸�
    memo = "�޸�"
    
    '���޹޴��ڿ��� ���۵Ǵ� ����ȳ����� ����, �̱���� �⺻������� ����
    emailSubject = ""
    
    '�������� ��������, �⺻�� - False
    '���ึ������ ���� ���ݰ�꼭�� �����ϴ� ���, ���꼼�� �ΰ��� �� �ֽ��ϴ�.
    '�������� ���ݰ�꼭�� �Ű��ؾ� �ϴ� ��� forceIssue ���� True�� �����Ͽ� ����(Issue API)�� ȣ���� �� �ֽ��ϴ�.
    forceIssue = False
    
    Set Response = TaxinvoiceService.Issue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, emailSubject, forceIssue)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message + vbCrLf + "����û���ι�ȣ : " + Response.ntsConfirmNum)
End Sub

'=========================================================================
' �����ڰ� ���޹޴��ڿ��� ������ ��û ���� ���ݰ�꼭�� ������ �ź��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#Refuse
'=========================================================================
Private Sub btnRefuse_sub_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim memo As String
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    '�޸�
    memo = "��)���� ��û �ź� �޸�"
    
    Set Response = TaxinvoiceService.Refuse(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����û ���� ���� "����Ϸ�" ������ ���ݰ�꼭�� "�������"�ϰ� ����û ���� ��󿡼� �����մϴ�.
' - ����(Delete API) �Լ��� ȣ���Ͽ� "�������" ������ ���ڼ��ݰ�꼭�� �����ϸ�, ������ȣ ������ �����մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#CancelIssue
'=========================================================================
Private Sub btnCancelIssue_rev_sub_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim memo As String
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    '�޸�
    memo = "������� �޸�"
    
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �����ڰ� ��û���� ������ ���ݰ�꼭�� �����ϱ� ��, ���޹޴��ڰ� �������û�� ����մϴ�.
' - �Լ� ȣ��� ���� ���� "���"�� ����ǰ�, �ش� ������ ���ݰ�꼭�� �����ڿ� ���� ���� �� �� �����ϴ�.
' - [���]�� ���ݰ�꼭�� ������ȣ�� �����ϱ� ���ؼ��� ���� (Delete API) �Լ��� ȣ���ؾ� �մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#CancelRequest
'=========================================================================
Private Sub btnRequestCancel_sub_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim memo As String
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    '�޸�
    memo = "��)���� ��û ��� �޸�"
    
    Set Response = TaxinvoiceService.CancelRequest(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���� ������ ������ ���ݰ�꼭�� �����մϴ�.
' - ���� ������ ����: "�ӽ�����", "�������", "������ź�", "���������", "���۽���"
' - ����ó���� ���ݰ�꼭�� ������ȣ�� ������ �����մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#Delete
'=========================================================================
Private Sub btnDelete_rev_sub_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select

    Set Response = TaxinvoiceService.Delete(txtCorpNum.Text, KeyType, txtMgtKey.Text)

    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �ۼ��� ���ݰ�꼭 �����͸� �˺��� �����մϴ�.
' - "�ӽ�����" ������ ���ݰ�꼭�� ����(Issue) �Լ��� ȣ���Ͽ� "����Ϸ�" ó���� ��쿡�� ����û���� ���۵˴ϴ�.
' - ������ �� �ӽ�����(Register)�� ����(Issue)�� �ѹ��� ȣ��� ó���ϴ� ��ù���(RegistIssue API) ���μ��� ������ �����մϴ�.
' - ������ �� �ӽ�����(Register)�� �������û(Request)�� �ѹ��� ȣ��� ó���ϴ� ��ÿ�û(RegistRequest API) ���μ��� ������ �����մϴ�.
' - ���ݰ�꼭 ����÷�� ����� �����ϴ� ���, �ӽ�����(Register API) -> ����÷��(AttachFile API) -> ����(Issue API) �Լ��� ���ʷ� ȣ���մϴ�.
' - ������ ���ݰ�꼭�� �����ϴ� ���, ��ü 'Taxinvoice'�� ���� 'chargeDirection' ���� ���� ���� ��ü�� ������ �� �ֽ��ϴ�.
'   �� ������ : ������ ���� , ������ : ���޹޴��� ����
' - �ӽ������ ���ݰ�꼭�� �˺� ����Ʈ '�ӽù�����'���� Ȯ�� �����մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#Register
'=========================================================================
Private Sub btnRegister_rev_Click()
    Dim Taxinvoice As New PBTaxinvoice
    
    '[�ʼ�] �ۼ�����, ǥ������ (yyyyMMdd)
    Taxinvoice.writeDate = "20220101"
    
    '[�ʼ�] ��������, [������, ������, ����Ź] �� ����
    Taxinvoice.issueType = "������"
    
    '[�ʼ�] {������, ������} �� ����, '������'�� ������ ���μ��������� �̿밡��
    '- ������(������ ����), ������(���޹޴��� ����)
    Taxinvoice.chargeDirection = "������"
    
    '[�ʼ�] ����/û��, [����, û��, ����] �� ����
    Taxinvoice.purposeType = "����"
        
    '[�ʼ�] ��������, [����, ����, �鼼] �� ����
    Taxinvoice.taxType = "����"
    
    
    '=========================================================================
    '                              ������ ����
    '=========================================================================
        
    '[�ʼ�] ������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    Taxinvoice.invoicerCorpNum = "8888888888"
    
    '[�ʼ�] ������ ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    Taxinvoice.invoicerTaxRegID = ""
    
    '[�ʼ�] ������ ��ȣ
    Taxinvoice.invoicerCorpName = "������ ��ȣ"
    
    '[�ʼ�] ������ ������ȣ, �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    Taxinvoice.invoicerMgtKey = ""
    
    '[�ʼ�] ������ ��ǥ�� ����
    Taxinvoice.invoicerCEOName = "������ ��ǥ�� ����"
    
    '������ �ּ�
    Taxinvoice.invoicerAddr = "������ �ּ�"
    
    '������ ����
    Taxinvoice.invoicerBizType = "������ ����,����2"
    
    '������ ����
    Taxinvoice.invoicerBizClass = "������ ����"
    
    '������ ����ڸ�
    Taxinvoice.invoicerContactName = "������ ����ڸ�"
    
    '������ ����� �����ּ�
    Taxinvoice.invoicerEmail = "test@test.com"
    
    '������ ����� ����ó
    Taxinvoice.invoicerTEL = "070-7070-0707"
    
    '������ ����� �޴�����ȣ
    Taxinvoice.invoicerHP = "010-000-2222"
    
    '������� ���޹޴��ڿ��� ����ȳ����� ���ۿ���
    '- �ȳ����� ���۱�� �̿�� ����Ʈ�� �����˴ϴ�.
    Taxinvoice.invoicerSMSSendYN = False
    
    
    '=========================================================================
    '                            ���޹޴��� ����
    '=========================================================================
        
    '[�ʼ�] ���޹޴��� ����, [�����, ����, �ܱ���] �� ����
    Taxinvoice.invoiceeType = "�����"
    
    '[�ʼ�] ���޹޴��� ����ڹ�ȣ, '-' ���� 10�ڸ�
    Taxinvoice.invoiceeCorpNum = txtCorpNum.Text
    
    '[�ʼ�] ���޹޴��� ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    Taxinvoice.invoiceeTaxRegID = ""
    
    '[�ʼ�] �����ڹ޴��� ��ȣ
    Taxinvoice.invoiceeCorpName = "���޹޴��� ��ȣ"
    
    '[������� �ʼ�] ���޹޴��� ������ȣ, �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    Taxinvoice.invoiceeMgtKey = txtMgtKey.Text
    
    '[�ʼ�] ���޹޴��� ��ǥ�� ����
    Taxinvoice.invoiceeCEOName = "���޹޴��� ��ǥ�� ����"
    
    '���޹޴��� �ּ�
    Taxinvoice.invoiceeAddr = "���޹޴��� �ּ�"
    
    '���޹޴��� ����
    Taxinvoice.invoiceeBizClass = "���޹޴��� ����"
    
    '���޹޴��� ����
    Taxinvoice.invoiceeBizType = "���޹޴��� ����"
    
    '���޹޴��� ����ڸ�
    Taxinvoice.invoiceeContactName1 = "���޹޴��� ����ڸ�"
    
    '���޹޴��� ����� �����ּ�
    '�˺� ����ȯ�濡�� �׽�Ʈ�ϴ� ��쿡�� �ȳ� ������ ���۵ǹǷ�,
    '���� �ŷ�ó�� �����ּҰ� ������� �ʵ��� ����
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    '���޹޴��� ����� ����ó
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    
    '���޹޴��� ����� �޴�����ȣ
    Taxinvoice.invoiceeHP1 = "010-111-222"
    
    ' ������ ��û�� �˸����� ���ۿ��� (�����࿡���� ��밡��)
    ' - ������ ����� �޴�����ȣ(invoicerHP)�� ����
    ' - ���۽� ����Ʈ�� �����Ǹ� ���۽����ϴ� ��� ����Ʈ ȯ��ó��
    Taxinvoice.invoiceeSMSSendYN = False
            
    
    '=========================================================================
    '                            ���ݰ�꼭 ����
    '=========================================================================
    
    '[�ʼ�] ���ް��� �հ�
    Taxinvoice.supplyCostTotal = "200000"
    
    '[�ʼ�] ���� �հ�
    Taxinvoice.taxTotal = "20000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + �����հ�
    Taxinvoice.totalAmount = "220000"
    
    '���� �� '�Ϸù�ȣ' �׸�
    Taxinvoice.serialNum = "123"
    
    '���� �� '��' �׸�, �ִ밪 32767
    '�̱���� Taxinvoice.kwon = ""
    Taxinvoice.kwon = "1"
    
    '���� �� 'ȣ' �׸�, �ִ밪 32767
    '�̱���� Taxinvoice.kwon = ""
    Taxinvoice.ho = "1"
    
    '���� �� '����' �׸�
    Taxinvoice.cash = ""
    
    '���� �� '��ǥ' �׸�
    Taxinvoice.chkBill = ""
    
    '���� �� '����' �׸�
    Taxinvoice.note = ""
    
    '���� �� '�ܻ�̼���' �׸�
    Taxinvoice.credit = ""
    
    '���� �� '���'�׸�
    Taxinvoice.remark1 = "���1"
    Taxinvoice.remark2 = "���2"
    Taxinvoice.remark3 = "���3"
    
    '����ڵ���� �̹��� ÷�ο���
    Taxinvoice.businessLicenseYN = False
    
    '����纻 �̹��� ÷�ο���
    Taxinvoice.bankBookYN = False
    

    '=========================================================================
    '         �������ݰ�꼭 ���� (�������ݰ�꼭 �ۼ��ÿ��� ����
    ' - �������ݰ�꼭 ���� ������ �����Ŵ��� �Ǵ� ���߰��̵� ��ũ ����
    ' - [����] �������ݰ�꼭 �ۼ���� �ȳ� - https://developers.popbill.com/guide/taxinvoice/vb/introduction/modified-taxinvoice
    '========================================================================='
    
    ' ���������ڵ�, ���������� ���� 1~6�� ���ñ���
    Taxinvoice.modifyCode = ""
    
    ' �������ݰ�꼭 ����û���ι�ȣ ����
    Taxinvoice.orgNTSConfirmNum = ""
        
    
    '=========================================================================
    '             ���׸�(ǰ��) ���� > �迭�� 99������ ���� ����
    '=========================================================================
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20220101"   '�ŷ�����  yyyyMMdd
    newDetail.itemName = "ǰ��"         'ǰ���
    newDetail.spec = "�԰�"             '�԰�
    newDetail.qty = "1"                 '����
    newDetail.unitCost = "100000"       '�ܰ�
    newDetail.supplyCost = "100000"     '���ް���
    newDetail.tax = "10000"             '����
    newDetail.remark = "���"           '���
    
    Taxinvoice.detailList.Add newDetail
    
    Dim newDetail2 As New PBTIDetail
    newDetail2.serialNum = 2             '�Ϸù�ȣ 1���� ���� ����
    newDetail2.purchaseDT = "20220101"   '�ŷ�����  yyyyMMdd
    newDetail2.itemName = "ǰ��2"        'ǰ���
    newDetail2.spec = "�԰�"             '�԰�
    newDetail2.qty = "1"                 '����
    newDetail2.unitCost = "100000"       '�ܰ�
    newDetail2.supplyCost = "100000"     '���ް���
    newDetail2.tax = "10000"             '����
    newDetail2.remark = "���"           '���
    
    Taxinvoice.detailList.Add newDetail2

    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.Register(txtCorpNum.Text, Taxinvoice, False)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' "�ӽ�����" ������ ���ݰ�꼭�� �����մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#Update
'=========================================================================
Private Sub btnUpdate_rev_Click()
    Dim KeyType As MgtKeyType
    
    '���ݰ�꼭 ��������, SELL-����, BUY-����, TRUSTEE-����Ź
    KeyType = BUY
    
    Dim Taxinvoice As New PBTaxinvoice
    
    '[�ʼ�] �ۼ�����, ǥ������ (yyyyMMdd)
    Taxinvoice.writeDate = "20220101"
    
    '[�ʼ�] ��������, [������, ������, ����Ź] �� ����
    Taxinvoice.issueType = "������"
    
    '[�ʼ�] {������, ������} �� ����, '������'�� ������ ���μ��������� �̿밡��
    '- ������(������ ����), ������(���޹޴��� ����)
    Taxinvoice.chargeDirection = "������"
    
    '[�ʼ�] ����/û��, [����, û��, ����] �� ����
    Taxinvoice.purposeType = "����"
        
    '[�ʼ�] ��������, [����, ����, �鼼] �� ����
    Taxinvoice.taxType = "����"
    
    
    '=========================================================================
    '                              ������ ����
    '=========================================================================
        
    '[�ʼ�] ������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    Taxinvoice.invoicerCorpNum = "8888888888"
    
    '[�ʼ�] ������ ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    Taxinvoice.invoicerTaxRegID = ""
    
    '[�ʼ�] ������ ��ȣ
    Taxinvoice.invoicerCorpName = "������ ��ȣ_����"
    
    '[�ʼ�] ������ ������ȣ, �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    Taxinvoice.invoicerMgtKey = ""
    
    '[�ʼ�] ������ ��ǥ�� ����
    Taxinvoice.invoicerCEOName = "������ ��ǥ�� ����_����"
    
    '������ �ּ�
    Taxinvoice.invoicerAddr = "������ �ּ�"
    
    '������ ����
    Taxinvoice.invoicerBizType = "������ ����,����2"
    
    '������ ����
    Taxinvoice.invoicerBizClass = "������ ����"
    
    '������ ����ڸ�
    Taxinvoice.invoicerContactName = "������ ����ڸ�"
    
    '������ ����� �����ּ�
    Taxinvoice.invoicerEmail = "test@test.com"
    
    '������ ����� ����ó
    Taxinvoice.invoicerTEL = "070-7070-0707"
    
    '������ ����� �޴�����ȣ
    Taxinvoice.invoicerHP = "010-000-2222"
    
    '=========================================================================
    '                            ���޹޴��� ����
    '=========================================================================
        
    '[�ʼ�] ���޹޴��� ����, [�����, ����, �ܱ���] �� ����
    Taxinvoice.invoiceeType = "�����"
    
    '[�ʼ�] ���޹޴��� ����ڹ�ȣ, '-' ���� 10�ڸ�
    Taxinvoice.invoiceeCorpNum = "1234567890"
    
    '[�ʼ�] ���޹޴��� ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    Taxinvoice.invoiceeTaxRegID = ""
    
    '[�ʼ�] �����ڹ޴��� ��ȣ
    Taxinvoice.invoiceeCorpName = "���޹޴��� ��ȣ"
    
    '[������� �ʼ�] ���޹޴��� ������ȣ, �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    Taxinvoice.invoiceeMgtKey = txtMgtKey.Text
    
    '[�ʼ�] ���޹޴��� ��ǥ�� ����
    Taxinvoice.invoiceeCEOName = "���޹޴��� ��ǥ�� ����"
    
    '���޹޴��� �ּ�
    Taxinvoice.invoiceeAddr = "���޹޴��� �ּ�"
    
    '���޹޴��� ����
    Taxinvoice.invoiceeBizClass = "���޹޴��� ����"
    
    '���޹޴��� ����
    Taxinvoice.invoiceeBizType = "���޹޴��� ����"
    
    '���޹޴��� ����ڸ�
    Taxinvoice.invoiceeContactName1 = "���޹޴��� ����ڸ�"
    
    '���޹޴��� ����� �����ּ�
    '�˺� ����ȯ�濡�� �׽�Ʈ�ϴ� ��쿡�� �ȳ� ������ ���۵ǹǷ�,
    '���� �ŷ�ó�� �����ּҰ� ������� �ʵ��� ����
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    '���޹޴��� ����� ����ó
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    
    '���޹޴��� ����� �޴�����ȣ
    Taxinvoice.invoiceeHP1 = "010-111-222"
    
    ' ������ ��û�� �˸����� ���ۿ��� (�����࿡���� ��밡��)
    ' - ������ ����� �޴�����ȣ(invoicerHP)�� ����
    ' - ���۽� ����Ʈ�� �����Ǹ� ���۽����ϴ� ��� ����Ʈ ȯ��ó��
    Taxinvoice.invoiceeSMSSendYN = False
            
    
    '=========================================================================
    '                            ���ݰ�꼭 ����
    '=========================================================================
    
    '[�ʼ�] ���ް��� �հ�
    Taxinvoice.supplyCostTotal = "200000"
    
    '[�ʼ�] ���� �հ�
    Taxinvoice.taxTotal = "20000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + �����հ�
    Taxinvoice.totalAmount = "220000"
    
    '���� �� '�Ϸù�ȣ' �׸�
    Taxinvoice.serialNum = "123"
    
    '���� �� '��' �׸�, �ִ밪 32767
    '�̱���� Taxinvoice.kwon = ""
    Taxinvoice.kwon = "1"
    
    '���� �� 'ȣ' �׸�, �ִ밪 32767
    '�̱���� Taxinvoice.kwon = ""
    Taxinvoice.ho = "1"
    
    '���� �� '����' �׸�
    Taxinvoice.cash = ""
    
    '���� �� '��ǥ' �׸�
    Taxinvoice.chkBill = ""
    
    '���� �� '����' �׸�
    Taxinvoice.note = ""
    
    '���� �� '�ܻ�̼���' �׸�
    Taxinvoice.credit = ""
    
    '���� �� '���'�׸�
    Taxinvoice.remark1 = "���1"
    Taxinvoice.remark2 = "���2"
    Taxinvoice.remark3 = "���3"
    
    '����ڵ���� �̹��� ÷�ο���
    Taxinvoice.businessLicenseYN = False
    
    '����纻 �̹��� ÷�ο���
    Taxinvoice.bankBookYN = False
    

    '=========================================================================
    '         �������ݰ�꼭 ���� (�������ݰ�꼭 �ۼ��ÿ��� ����
    ' - �������ݰ�꼭 ���� ������ �����Ŵ��� �Ǵ� ���߰��̵� ��ũ ����
    ' - [����] �������ݰ�꼭 �ۼ���� �ȳ� - https://developers.popbill.com/guide/taxinvoice/vb/introduction/modified-taxinvoice
    '========================================================================='
    
    ' ���������ڵ�, ���������� ���� 1~6�� ���ñ���
    Taxinvoice.modifyCode = ""
    
    ' �������ݰ�꼭 ����û���ι�ȣ ����
    Taxinvoice.orgNTSConfirmNum = ""
        
    
    '=========================================================================
    '             ���׸�(ǰ��) ���� > �迭�� 99������ ���� ����
    '=========================================================================
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20220101"   '�ŷ�����  yyyyMMdd
    newDetail.itemName = "ǰ��"         'ǰ���
    newDetail.spec = "�԰�"             '�԰�
    newDetail.qty = "1"                 '����
    newDetail.unitCost = "100000"       '�ܰ�
    newDetail.supplyCost = "100000"     '���ް���
    newDetail.tax = "10000"             '����
    newDetail.remark = "���"           '���
    
    Taxinvoice.detailList.Add newDetail
    
    Dim newDetail2 As New PBTIDetail
    newDetail2.serialNum = 2             '�Ϸù�ȣ 1���� ���� ����
    newDetail2.purchaseDT = "20220101"   '�ŷ�����  yyyyMMdd
    newDetail2.itemName = "ǰ��2"        'ǰ���
    newDetail2.spec = "�԰�"             '�԰�
    newDetail2.qty = "1"                 '����
    newDetail2.unitCost = "100000"       '�ܰ�
    newDetail2.supplyCost = "100000"     '���ް���
    newDetail2.tax = "10000"             '����
    newDetail2.remark = "���"           '���
    
    Taxinvoice.detailList.Add newDetail2
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.Update(txtCorpNum.Text, KeyType, txtMgtKey.Text, Taxinvoice, False)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���޹޴��ڰ� ����� ������ ���ݰ�꼭�� �����ڿ��� �ۺ��Ͽ� ���� ��û�մϴ�.
' - ������ ���ݰ�꼭 ���μ����� �����ϱ� ���ؼ��� ������/���޹޴��ڰ� ��� �˺��� ȸ���̿��� �մϴ�.
' - ������ ��û�� ���ݰ�꼭�� "(��)������" �����̸�, �����ڰ� �˺� ����Ʈ �Ǵ� �Լ��� ȣ���Ͽ� ������ ��쿡�� ����û���� ���۵˴ϴ�.
' - �����ڴ� �˺� ����Ʈ�� "���� ���� �����"���� ������ ������ ������ ���ݰ�꼭�� Ȯ���� �� �ֽ��ϴ�.
' - ������ ��û�� �����ڿ��� ������ ��û ������ �߼۵˴ϴ�.
' - �����ڰ� ������ ���ݰ�꼭 ����� ����Ʈ�� ���ݵ˴ϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#Request
'=========================================================================
Private Sub btnRequest_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim memo As String
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    '�޸�
    memo = "������ ��û �޸�"
    
    
    Set Response = TaxinvoiceService.Request(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'======================================================================================================================
' "�ӽ�����" �Ǵ� "(��)������" ������ ���ݰ�꼭�� ����(���ڼ���)�ϸ�, "����Ϸ�" ���·� ó���մϴ�.
' - ���ݰ�꼭 ����û ������å [https://developers.popbill.com/guide/taxinvoice/vb/introduction/policy-of-send-to-nts]
' - "����Ϸ�" �� ���ڼ��ݰ�꼭�� ����û ���� ������ �������(CancelIssue API) �Լ��� ����û �Ű� ��󿡼� ������ �� �ֽ��ϴ�.
' - ���ݰ�꼭 ������ ���ؼ� �������� �������� �˺� ���������� ������� �Ǿ�� �մϴ�.
'   �� ����Ź������ ���, ��Ź���� ������ ����� �ʿ��մϴ�.
' - ���ݰ�꼭 ���� �� ���޹޴��ڿ��� ���� ������ �߼۵˴ϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#Issue
'======================================================================================================================
Private Sub btnIssue_rev_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim memo As String
    Dim emailSubject As String
    Dim forceIssue As Boolean
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    '�޸�
    memo = "������ ���ݰ�꼭 ����"
    
    '���޹޴��ڿ��� ���۵Ǵ� ����ȳ����� ����, �̱���� �⺻������� ����
    emailSubject = ""
    
    '�������� ��������, �⺻�� - False
    '���ึ������ ���� ���ݰ�꼭�� �����ϴ� ���, ���꼼�� �ΰ��� �� �ֽ��ϴ�.
    '�������� ���ݰ�꼭�� �Ű��ؾ� �ϴ� ��� forceIssue ���� True�� �����Ͽ� ����(Issue API)�� ȣ���� �� �ֽ��ϴ�.
    forceIssue = False
    
    Set Response = TaxinvoiceService.Issue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, emailSubject, forceIssue)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message + vbCrLf + "����û���ι�ȣ : " + Response.ntsConfirmNum)
End Sub

'=========================================================================
' �����ڰ� ���޹޴��ڿ��� ������ ��û ���� ���ݰ�꼭�� ������ �ź��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#Refuse
'=========================================================================
Private Sub btnRefuse_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim memo As String
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    '�޸�
    memo = "������ ��û �ź� �޸�"
    
    Set Response = TaxinvoiceService.Refuse(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
'����û ���� ���� "����Ϸ�" ������ ���ݰ�꼭�� "�������"�ϰ� ����û ���� ��󿡼� �����մϴ�.
' - ����(Delete API) �Լ��� ȣ���Ͽ� "�������" ������ ���ڼ��ݰ�꼭�� �����ϸ�, ������ȣ ������ �����մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#CancelIssue
'=========================================================================
Private Sub btnCancelIssue_rev_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim memo As String
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    '�޸�
    memo = "������� �޸�"
    
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �����ڰ� ��û���� ������ ���ݰ�꼭�� �����ϱ� ��, ���޹޴��ڰ� �������û�� ����մϴ�.
' - �Լ� ȣ��� ���� ���� "���"�� ����ǰ�, �ش� ������ ���ݰ�꼭�� �����ڿ� ���� ���� �� �� �����ϴ�.
' - [���]�� ���ݰ�꼭�� ������ȣ�� �����ϱ� ���ؼ��� ���� (Delete API) �Լ��� ȣ���ؾ� �մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#CancelRequest
'=========================================================================
Private Sub btnRequestCancel_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim memo As String
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    '�޸�
    memo = "��)���� ��û ��� �޸�"
    
    Set Response = TaxinvoiceService.CancelRequest(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' 1���� ������ ������ ���ݰ�꼭�� �����մϴ�.
' - ���� ������ ����: "�ӽ�����", "�������", "������ź�", "���������", "���۽���"
' - ����ó���� ���ݰ�꼭�� ������ȣ�� ������ �����մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/issue#Delete
'=========================================================================
Private Sub btnDelete_rev_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select

    Set Response = TaxinvoiceService.Delete(txtCorpNum.Text, KeyType, txtMgtKey.Text)

    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' "�ӽ�����" ������ ���ݰ�꼭�� 1���� ������ ÷���մϴ�. (�ִ� 5��)
' - https://developers.popbill.com/reference/taxinvoice/vb/api/etc#AttachFile
'=========================================================================
Private Sub btnAttachFile_Click()
    Dim FilePath As String
    CommonDialog1.FileName = ""

    CommonDialog1.ShowOpen

    FilePath = CommonDialog1.FileName

    If FilePath = "" Then Exit Sub

    Dim Response As PBResponse
    Dim KeyType As MgtKeyType

    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select

    Set Response = TaxinvoiceService.AttachFile(txtCorpNum.Text, KeyType, txtMgtKey.Text, FilePath)

    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If

    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ݰ�꼭�� ÷�ε� ���ϸ���� Ȯ���մϴ�.
' - �����׸� �� ���Ͼ��̵�(AttachedFile) �׸��� ÷������ ����(DeleteFile API) �Լ� ȣ�� �� �̿��� �� �ֽ��ϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/etc#GetFiles
'=========================================================================
Private Sub btnGetFiles_Click()
    Dim resultList As Collection
    Dim KeyType As MgtKeyType
    Dim tmp As String
    Dim file As PBAttachFile
    
    Set resultList = TaxinvoiceService.GetFiles(txtCorpNum.Text, KeyType, txtMgtKey.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "serialNum(�Ϸù�ȣ) | attachedfile(���Ͼ��̵�) | displayName(÷�����ϸ�) |  RegDT(÷���Ͻ�)" + vbCrLf
    
    For Each file In resultList
        tmp = tmp + CStr(file.serialNum) + " | " + file.AttachedFile + " | " + file.DisplayName + " | " + file.regDT + vbCrLf
        txtFileID.Text = file.AttachedFile
        
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' "�ӽ�����" ������ ���ݰ�꼭�� ÷�ε� 1���� ������ �����մϴ�.
' - ���� �ĺ��� ���� ÷�� �� �ο��Ǵ� 'FileID'�� ÷������ ��� Ȯ��(GetFiles API) �Լ��� ȣ���Ͽ� Ȯ���մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/etc#DeleteFile
'=========================================================================
Private Sub btnDeleteFile_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.DeleteFile(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtFileID.Text)
            
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ݰ�꼭 1���� ���� �� ��������� Ȯ���մϴ�.
' - ���ϰ� 'PBTIInfo'�� ���� 'stateCode'�� ���� ���ݰ�꼭�� �����ڵ带 Ȯ���մϴ�.
' - ���ݰ�꼭 �����ڵ� [https://developers.popbill.com/reference/taxinvoice/vb/response-code#state-code]
' - https://developers.popbill.com/reference/taxinvoice/vb/api/info#GetInfo
'=========================================================================
Private Sub btnGetInfo_Click()
    Dim tiInfo As PBTIInfo
    Dim KeyType As MgtKeyType
    Dim tmp As String
   
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    Set tiInfo = TaxinvoiceService.GetInfo(txtCorpNum.Text, KeyType, txtMgtKey.Text)
     
    If tiInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "itemKey (�˺���ȣ) : " + tiInfo.itemKey + vbCrLf
    tmp = tmp + "taxType (��������) : " + tiInfo.taxType + vbCrLf
    tmp = tmp + "writeDate (�ۼ�����) : " + tiInfo.writeDate + vbCrLf
    tmp = tmp + "regDT (�ӽ����� ����) : " + tiInfo.regDT + vbCrLf
    tmp = tmp + "issueType (��������) : " + tiInfo.issueType + vbCrLf
    tmp = tmp + "supplyCostTotal (���ް��� �հ�) : " + tiInfo.supplyCostTotal + vbCrLf
    tmp = tmp + "taxTotal (���� �հ�) : " + tiInfo.taxTotal + vbCrLf
    tmp = tmp + "purposeType (����/û��) : " + tiInfo.purposeType + vbCrLf
    tmp = tmp + "issueDT (�����Ͻ�) : " + tiInfo.issueDT + vbCrLf
    tmp = tmp + "stateDT (���� �����Ͻ�) : " + tiInfo.stateDT + vbCrLf
    tmp = tmp + "lateIssueYN (�������� ����) : " + CStr(tiInfo.lateIssueYN) + vbCrLf
    tmp = tmp + "openYN (��������) : " + CStr(tiInfo.openYN) + vbCrLf
    tmp = tmp + "openDT (�����Ͻ�) : " + tiInfo.openDT + vbCrLf
    tmp = tmp + "stateCode (�����ڵ�) : " + CStr(tiInfo.stateCode) + vbCrLf
    tmp = tmp + "stateMemo (���¸޸�) : " + tiInfo.stateMemo + vbCrLf
    tmp = tmp + "ntsresult (����û ���۰��) : " + tiInfo.ntsresult + vbCrLf
    tmp = tmp + "ntsconfirmNum (����û���ι�ȣ) : " + tiInfo.ntsConfirmNum + vbCrLf
    tmp = tmp + "ntssendDT (����û �����Ͻ�) : " + tiInfo.ntssendDT + vbCrLf
    tmp = tmp + "ntsresultDT (����û ��� �����Ͻ�) : " + tiInfo.ntsresultDT + vbCrLf
    tmp = tmp + "ntssendErrCode (���۽��� �����ڵ�) : " + tiInfo.ntssendErrCode + vbCrLf
    tmp = tmp + "modifyCode (���� �����ڵ�) : " + tiInfo.modifyCode + vbCrLf
    tmp = tmp + "interOPYN (�������� ����) : " + CStr(tiInfo.interOPYN) + vbCrLf
    tmp = tmp + "invoicerCorpName (������ ��ȣ) : " + tiInfo.invoicerCorpName + vbCrLf
    tmp = tmp + "invoicerCorpNum (������ ����ڹ�ȣ) : " + tiInfo.invoicerCorpNum + vbCrLf
    tmp = tmp + "invoicerMgtKey (������ ������ȣ) : " + tiInfo.invoicerMgtKey + vbCrLf
    tmp = tmp + "invoicerPrintYN (������ �μ⿩��) : " + CStr(tiInfo.invoicerPrintYN) + vbCrLf
    tmp = tmp + "invoiceeCorpName (���޹޴��� ��ȣ) : " + tiInfo.invoiceeCorpName + vbCrLf
    tmp = tmp + "invoiceeCorpNum (���޹޴��� ����ڹ�ȣ) : " + tiInfo.invoiceeCorpNum + vbCrLf
    tmp = tmp + "invoiceeMgtKey (���޹޴��� ������ȣ) : " + tiInfo.invoiceeMgtKey + vbCrLf
    tmp = tmp + "invoiceePrintYN (���޹޴��� �μ⿩��) : " + CStr(tiInfo.invoiceePrintYN) + vbCrLf
    tmp = tmp + "closeDownState (���޹޴��� ���������) : " + CStr(tiInfo.closeDownState) + vbCrLf
    tmp = tmp + "closeDownStateDate (���޹޴��� ��������� : " + tiInfo.closeDownStateDate + vbCrLf
    tmp = tmp + "trusteeCorpName (��Ź�� ��ȣ) : " + tiInfo.trusteeCorpName + vbCrLf
    tmp = tmp + "trusteeCorpNum (��Ź�� ����ڹ�ȣ) : " + tiInfo.trusteeCorpNum + vbCrLf
    tmp = tmp + "trusteeMgtKey (��Ź�� ������ȣ) : " + tiInfo.trusteeMgtKey + vbCrLf
    tmp = tmp + "trusteePrintYN (��Ź�� �μ⿩��) : " + CStr(tiInfo.trusteePrintYN) + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' �ټ����� ���ݰ�꼭 ���� �� ��� ������ Ȯ���մϴ�. (1ȸ ȣ�� �� �ִ� 1,000�� Ȯ�� ����)
' ���ϰ� 'PBTIInfo'�� ���� 'stateCode'�� ���� ���ݰ�꼭�� �����ڵ带 Ȯ���մϴ�.
' ���ݰ�꼭 �����ڵ� [https://developers.popbill.com/reference/taxinvoice/vb/response-code#state-code]
' - https://developers.popbill.com/reference/taxinvoice/vb/api/info#GetInfos
'=========================================================================
Private Sub btnGetInfos_Click()
    Dim resultList As Collection
    Dim KeyType As MgtKeyType
    Dim KeyList As New Collection
    Dim tmp As String
    Dim info As PBTIInfo
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    '���ݰ�꼭 ������ȣ �迭, �ִ� 1000��
    KeyList.Add "20220101-01"
    KeyList.Add "20220101-02"
    KeyList.Add "20220101-03"
    KeyList.Add "20220101-04"
    
    Set resultList = TaxinvoiceService.GetInfos(txtCorpNum.Text, KeyType, KeyList)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "itemKey(�˺���ȣ) | taxType (��������) | writeDate (�ۼ�����) | regDT (�ӽ����� �Ͻ�) | issueType (��������) | supplyCostTotal (���ް��� �հ�) | " + vbCrLf
    tmp = tmp + "taxTotal (���� �հ�) | purposeType (����/û��) |issueDT (�����Ͻ�) | lateIssueYN (�������� ����) | openYN (���� ����) | openDT (���� �Ͻ�) | " + vbCrLf
    tmp = tmp + "stateMemo (���¸޸�) | stateCode (�����ڵ�) | ntsconfirmNum (����û���ι�ȣ) | ntsresult (����û ���۰��) | ntssendDT (����û �����Ͻ�) | " + vbCrLf
    tmp = tmp + "ntsresultDT (����û ��� �����Ͻ�) | ntssendErrCode (���л��� �����ڵ�) | modifyCode (���� �����ڵ�) | interOPYN (�������� ����) | invoicerCorpName (������ ��ȣ) | " + vbCrLf
    tmp = tmp + "invoicerCorpNum (������ ����ڹ�ȣ) | invoicerMgtKey (������ ������ȣ) | invoicerPrintYN (������ �μ⿩��) | invoiceeCorpName (���޹޴��� ��ȣ) | " + vbCrLf
    tmp = tmp + "invoiceeCorpNum (���޹޴��� ����ڹ�ȣ) | invoiceeMgtKey(���޹޴��� ������ȣ) | invoiceePrintYN(���޹޴��� �μ⿩��) | closeDownState(���޹޴��� ���������) | " + vbCrLf
    tmp = tmp + "closeDownStateDate(���޹޴��� ���������) | trusteeCorpName (��Ź�� ��ȣ) | trusteeCorpNum (��Ź�� ����ڹ�ȣ) | trusteeMgtKey(��Ź�� ������ȣ) | " + vbCrLf
    tmp = tmp + "trusteePrintYN(��Ź�� �μ⿩��) " + vbCrLf
    
    
    For Each info In resultList
        tmp = tmp + info.itemKey + " | " + info.taxType + " | " + info.writeDate + " | " + info.regDT + " | " + info.issueType + " | " + vbCrLf
        tmp = tmp + info.supplyCostTotal + " | " + info.taxTotal + " | " + info.purposeType + " | " + info.issueDT + " | " + vbCrLf
        tmp = tmp + info.stateDT + " | " + CStr(info.lateIssueYN) + " | " + CStr(info.openYN) + " | " + info.openDT + " | " + vbCrLf
        tmp = tmp + CStr(info.stateCode) + " | " + info.stateMemo + " | " + info.ntsresult + " | " + info.ntsConfirmNum + " | " + vbCrLf
        tmp = tmp + info.ntssendDT + " | " + info.ntsresultDT + " | " + info.ntssendErrCode + " | " + info.modifyCode + " | " + CStr(info.interOPYN) + " | " + vbCrLf
        tmp = tmp + info.invoicerCorpName + " | " + info.invoicerCorpNum + " | " + info.invoicerMgtKey + " | " + CStr(info.invoicerPrintYN) + " | " + vbCrLf
        tmp = tmp + info.invoiceeCorpName + " | " + info.invoiceeCorpNum + " | " + info.invoiceeMgtKey + " | " + vbCrLf
        tmp = tmp + CStr(info.invoiceePrintYN) + " | " + CStr(info.closeDownState) + " | " + info.closeDownStateDate + " | " + vbCrLf
        tmp = tmp + info.trusteeCorpName + " | " + info.trusteeCorpNum + " | " + info.trusteeMgtKey + " | " + CStr(info.trusteePrintYN) + " | " + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ���ݰ�꼭 1���� �������� Ȯ���մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/info#GetDetailInfo
'=========================================================================
Private Sub btnGetDetailInfo_Click()
    Dim tiDetailInfo As PBTaxinvoice
    Dim detail As PBTIDetail
    Dim contact As PBTIContact
    Dim KeyType As MgtKeyType
    Dim tmp As String
    
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    Set tiDetailInfo = TaxinvoiceService.GetDetailInfo(txtCorpNum.Text, KeyType, txtMgtKey.Text)
     
    If tiDetailInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ntsconfirmNum (����û ���ι�ȣ) : " + tiDetailInfo.ntsConfirmNum + vbCrLf
    tmp = tmp + "issueType (��������) : " + tiDetailInfo.issueType + vbCrLf
    tmp = tmp + "taxType (��������) : " + tiDetailInfo.taxType + vbCrLf
    tmp = tmp + "chargeDirection (���ݹ���) : " + tiDetailInfo.chargeDirection + vbCrLf
    tmp = tmp + "serialNum (�Ϸù�ȣ) : " + tiDetailInfo.serialNum + vbCrLf
    tmp = tmp + "kwon (��) : " + tiDetailInfo.kwon + vbCrLf
    tmp = tmp + "ho (ȣ) : " + tiDetailInfo.ho + vbCrLf
    tmp = tmp + "writeDate (�ۼ�����) : " + tiDetailInfo.writeDate + vbCrLf
    tmp = tmp + "purposeType (����/û��) : " + tiDetailInfo.purposeType + vbCrLf
    tmp = tmp + "supplyCostTotal (���ް��� �հ�) : " + tiDetailInfo.supplyCostTotal + vbCrLf
    tmp = tmp + "taxtotal (���� �հ�) : " + tiDetailInfo.taxTotal + vbCrLf
    tmp = tmp + "totalAmount (�հ� �ݾ�) : " + tiDetailInfo.totalAmount + vbCrLf
    tmp = tmp + "cash (����) : " + tiDetailInfo.cash + vbCrLf
    tmp = tmp + "chkbill (��ǥ) : " + tiDetailInfo.chkBill + vbCrLf
    tmp = tmp + "credit (�ܻ�) : " + tiDetailInfo.credit + vbCrLf
    tmp = tmp + "note (����) : " + tiDetailInfo.note + vbCrLf
    tmp = tmp + "remark1 (���1) : " + tiDetailInfo.remark1 + vbCrLf
    tmp = tmp + "remark2 (���2) : " + tiDetailInfo.remark2 + vbCrLf
    tmp = tmp + "remark3 (���3) : " + tiDetailInfo.remark3 + vbCrLf
        
    tmp = tmp + "invoicerCorpNum (������ ����ڹ�ȣ) : " + tiDetailInfo.invoicerCorpNum + vbCrLf
    tmp = tmp + "invoicerMgtKey (������ ������ȣ) : " + tiDetailInfo.invoicerMgtKey + vbCrLf
    tmp = tmp + "invoicerTaxRegID (������ ������� �ĺ���ȣ) : " + tiDetailInfo.invoicerTaxRegID + vbCrLf
    tmp = tmp + "invoicerCorpName (������ ��ȣ) : " + tiDetailInfo.invoicerCorpName + vbCrLf
    tmp = tmp + "invoicerCEOName (������ ��ǥ�� ����) : " + tiDetailInfo.invoicerCEOName + vbCrLf
    tmp = tmp + "invoicerAddr (������ �ּ�) : " + tiDetailInfo.invoicerAddr + vbCrLf
    tmp = tmp + "invoicerBizClass (������ ����) : " + tiDetailInfo.invoicerBizClass + vbCrLf
    tmp = tmp + "invoicerBizType (������ ����) : " + tiDetailInfo.invoicerBizType + vbCrLf
    tmp = tmp + "invoicerContactName (������ ����ڸ�) : " + tiDetailInfo.invoicerContactName + vbCrLf
    tmp = tmp + "invoicerDeptName (������ ����� �μ���) : " + tiDetailInfo.invoicerDeptName + vbCrLf
    tmp = tmp + "invoicerTEL (������ ����� ����ó) : " + tiDetailInfo.invoicerTEL + vbCrLf
    tmp = tmp + "invoicerHP (������ ����� �޴�����ȣ) : " + tiDetailInfo.invoicerHP + vbCrLf
    tmp = tmp + "invoicerEmail (������ ����� ����) : " + tiDetailInfo.invoicerEmail + vbCrLf
    tmp = tmp + "invoicerSMSSendYN (����ȳ����� ���ۿ���) : " + CStr(tiDetailInfo.invoicerSMSSendYN) + vbCrLf + vbCrLf
    
    tmp = tmp + "invoiceeCorpNum (���޹޴��� ����ڹ�ȣ) : " + tiDetailInfo.invoiceeCorpNum + vbCrLf
    tmp = tmp + "invoiceeType (���޹޴��� ����) : " + tiDetailInfo.invoiceeType + vbCrLf
    tmp = tmp + "invoiceeMgtKey (���޹޴��� ������ȣ) : " + tiDetailInfo.invoiceeMgtKey + vbCrLf
    tmp = tmp + "invoiceeTaxRegID (���޹޴��� ������� �ĺ���ȣ) : " + tiDetailInfo.invoiceeTaxRegID + vbCrLf
    tmp = tmp + "invoiceeCorpName (���޹޴��� ��ȣ) : " + tiDetailInfo.invoiceeCorpName + vbCrLf
    tmp = tmp + "invoiceeCEOName (���޹޴��� ��ǥ�� ����) : " + tiDetailInfo.invoiceeCEOName + vbCrLf
    tmp = tmp + "invoiceeAddr (���޹޴��� �ּ�) : " + tiDetailInfo.invoiceeAddr + vbCrLf
    tmp = tmp + "invoiceeBizClass (���޹޴��� ����) : " + tiDetailInfo.invoiceeBizClass + vbCrLf
    tmp = tmp + "invoiceeBizType (���޹޴��� ����) : " + tiDetailInfo.invoiceeBizType + vbCrLf
    tmp = tmp + "invoiceeContactName1 (���޹޴��� ����ڸ�) : " + tiDetailInfo.invoiceeContactName1 + vbCrLf
    tmp = tmp + "invoiceeDeptName1 (���޹޴��� ����� �μ���) : " + tiDetailInfo.invoiceeDeptName1 + vbCrLf
    tmp = tmp + "invoiceeTEL1 (���޹޴��� ����� ����ó) : " + tiDetailInfo.invoiceeTEL1 + vbCrLf
    tmp = tmp + "invoiceeHP1 (���޹޴��� ����� �޴�����ȣ) : " + tiDetailInfo.invoiceeHP1 + vbCrLf
    tmp = tmp + "invoiceeEmail1 (���޹޴��� ����� ����) : " + tiDetailInfo.invoiceeEmail1 + vbCrLf
    tmp = tmp + "closeDownState (���޹޴��� ���������) : " + CStr(tiDetailInfo.closeDownState) + vbCrLf
    tmp = tmp + "closeDownStateDate (���޹޴��� ���������) : " + tiDetailInfo.closeDownStateDate + vbCrLf + vbCrLf

    tmp = tmp + "modifyCode(�������� �ڵ�) : " + tiDetailInfo.modifyCode + vbCrLf
    tmp = tmp + "orgNTSConfirmNum(���� ���ݰ�꼭 ����û���ι�ȣ) : " + tiDetailInfo.orgNTSConfirmNum + vbCrLf
   
    If (tiDetailInfo.detailList Is Nothing) = False Then
        For Each detail In tiDetailInfo.detailList
            tmp = tmp + "serialNum (�Ϸù�ȣ) : " + CStr(detail.serialNum) + vbCrLf
            tmp = tmp + "purchaseDT (�ŷ�����) : " + detail.purchaseDT + vbCrLf
            tmp = tmp + "itemName (ǰ��) : " + detail.itemName + vbCrLf
            tmp = tmp + "spec (�԰�) : " + detail.spec + vbCrLf
            tmp = tmp + "qty (����) : " + detail.qty + vbCrLf
            tmp = tmp + "unitcost (�ܰ�) : " + detail.unitCost + vbCrLf
            tmp = tmp + "supplycost (���ް���) : " + detail.supplyCost + vbCrLf
            tmp = tmp + "tax (����) : " + detail.tax + vbCrLf
            tmp = tmp + "remark (���) : " + detail.remark + vbCrLf + vbCrLf
        Next
    End If
    
    If (tiDetailInfo.addContactList Is Nothing) = False Then
        For Each contact In tiDetailInfo.addContactList
            tmp = tmp + "serialNum (�Ϸù�ȣ) : " + CStr(contact.serialNum) + vbCrLf
            tmp = tmp + "contactName (����� ����) : " + contact.ContactName + vbCrLf
            tmp = tmp + "email (�̸����ּ�) : " + contact.email + vbCrLf + vbCrLf
        Next
    End If
    
    MsgBox tmp
End Sub

'=========================================================================
' ���ݰ�꼭 1���� �������� XML�� ��ȯ�մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/info#GetXML
'=========================================================================
Private Sub btnGetXML_Click()
    Dim result As PBTaxinvoiceXML
    Dim KeyType As MgtKeyType
    Dim tmp As String
    
    '���ݰ�꼭 ��������, SELL-����, BUY-����, TRUSTEE-����Ź
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    Set result = TaxinvoiceService.GetXML(txtCorpNum.Text, KeyType, txtMgtKey.Text)
     
    If result Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "code (�����ڵ�) : " + CStr(result.code) + vbCrLf
    tmp = tmp + "message (����޽���) : " + result.message + vbCrLf
    tmp = tmp + "retObject (���ڼ��ݰ�꼭 XML ���� ) : " + result.retObject + vbCrLf
    
    MsgBox tmp
    
End Sub

'=========================================================================
' �˻����ǿ� �ش��ϴ� ���ݰ�꼭�� ��ȸ�մϴ�. (��ȸ�Ⱓ ���� : �ִ� 6����)
' - https://developers.popbill.com/reference/taxinvoice/vb/api/info#Search
'=========================================================================
Private Sub btnSearch_Click()
    Dim tiSearchList As PBTISearchList
    Dim KeyType As MgtKeyType
    Dim DType As String
    Dim SDate As String
    Dim EDate As String
    Dim state As New Collection
    Dim TType As New Collection
    Dim taxType As New Collection
    Dim issueType As New Collection
    Dim regType As New Collection
    Dim closeDownState As New Collection
    Dim LateOnly As String
    Dim Page As Integer
    Dim PerPage As Integer
    Dim Order As String
    Dim TaxRegIDType As String
    Dim TaxRegID As String
    Dim TaxRegIDYN As String
    Dim QString As String
    Dim mgtKey As String
    Dim tmp As String
    Dim interOPYN As String
        
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    
    '[�ʼ�] ��������, R-����Ͻ� W-�ۼ����� I-�����Ͻ� �� ��1
    DType = "W"
    
    '[�ʼ�] ��������, yyyyMMdd
    SDate = "20220101"
    
    '[�ʼ�] ��������, yyyyMMdd
    EDate = "20220110"
    
    '���ۻ��°� �迭, �̱���� ��ü������ȸ, �������°� 3�ڸ����� �ۼ� 2,3��° ���ϵ�ī�� ����
    state.Add "3**"
    state.Add "6**"
    
    '�������� �迭, N-�Ϲ� M-���� �� ����, �̱���� ��ü��ȸ
    TType.Add "N"
    TType.Add "M"
    
    '�������� �迭, T-����, N-�鼼 Z-���� �� ����, �̱���� ��ü��ȸ
    taxType.Add "T"
    taxType.Add "N"
    taxType.Add "Z"
    
    '�������� �迭, N-������, R-������ T-����Ź
    issueType.Add "N"
    issueType.Add "R"
    issueType.Add "T"
    
    ' ������� �迭, P-�˺�, H-Ȩ�ý� �Ǵ� �ܺ�ASP
    regType.Add "P"
    regType.Add "H"
    
    '�������ȸ ���� �迭,  N-��Ȯ�� / 0-�̵�� / 1-����� / 2-��� / 3-�޾�
    closeDownState.Add "N"
    closeDownState.Add "0"
    closeDownState.Add "1"
    closeDownState.Add "2"
    closeDownState.Add "3"
    
    '�������� ����, 0-������� ��ȸ 1-�������� ��ȸ, ����ó���� ��ü��ȸ
    LateOnly = ""
    
    ' ���ڼ��ݰ�꼭 ������ȣ �Ǵ� ����û ���ι�ȣ �˻� ��ȸ, ����ó���� ��ü��ȸ
    mgtKey = ""
    
    '��������ȣ, �⺻�� ��1
    Page = 1
    
    '�������� �˻�����, �⺻�� 500, �ִ� 1000
    PerPage = 10
    
    '���Ĺ���, D-��������(�⺻��), A-��������
    Order = "D"
    
    '��������ȣ ���� S-������, B-���޹޴���, T-��Ź��
    TaxRegIDType = "S"
    
    '��������ȣ, �޸�(,)�� �����Ͽ� ���� ex) 0001,0002
    TaxRegID = ""
    
    '������� ����, ����-��ü��ȸ, 0-��������ȣ ���°�츸 ��ȸ, 1-��������ȣ ���� ��ȸ
    TaxRegIDYN = ""
    
    '�ŷ�ó ��ȸ, �ŷ�ó ��ȣ �Ǵ� �ŷ�ó ����ڵ�Ϲ�ȣ ��ȸ, ����ó���� ��ü��ȸ
    QString = ""
    
    '�������� ��ȸ ����, ����-��ü��ȸ, 0-�Ϲݹ��� ��ȸ, 1-�������� ��ȸ
    interOPYN = ""
    
    Set tiSearchList = TaxinvoiceService.Search(txtCorpNum.Text, KeyType, DType, SDate, EDate, state, _
                    TType, taxType, LateOnly, Page, PerPage, Order, TaxRegIDType, TaxRegID, TaxRegIDYN, QString, _
                    txtUserID.Text, interOPYN, issueType, regType, closeDownState, mgtKey)
     
    If tiSearchList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "code (�����ڵ�) : " + CStr(tiSearchList.code) + vbCrLf
    tmp = tmp + "total (�� �˻���� �Ǽ�) : " + CStr(tiSearchList.total) + vbCrLf
    tmp = tmp + "perPage (�������� �˻�����) : " + CStr(tiSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum (������ ��ȣ) : " + CStr(tiSearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (������ ����) : " + CStr(tiSearchList.pageCount) + vbCrLf
    tmp = tmp + "message (����޽���) : " + tiSearchList.message + vbCrLf + vbCrLf
    
    tmp = tmp + "itemKey(�˺���ȣ) |  taxType (��������) |  writeDate (�ۼ�����) |  regDT (�ӽ����� �Ͻ�) |  issueType (��������) |  supplyCostTotal (���ް��� �հ�) | " + _
         "taxTotal (���� �հ�) |  purposeType (����/û��) | issueDT (�����Ͻ�) | lateIssueYN (�������� ����) | openYN (���� ����) | openDT (���� �Ͻ�) | " + _
         "stateMemo (���¸޸�) | stateCode (�����ڵ�) | ntsconfirmNum (����û���ι�ȣ) | ntsresult (����û ���۰��) | ntssendDT (����û �����Ͻ�) | " + _
         "ntsresultDT (����û ��� �����Ͻ�) | ntssendErrCode (���۽��� �����ڵ�) | modifyCode (���� �����ڵ�) | interOPYN (�������� ����) | invoicerCorpName (������ ��ȣ) | " + _
         "invoicerCorpNum (������ ����ڹ�ȣ) | invoicerMgtKey (������ ������ȣ) | invoicerPrintYN (������ �μ⿩��) | invoiceeCorpName (���޹޴��� ��ȣ) | " + _
         "invoiceeCorpNum (���޹޴��� ����ڹ�ȣ) | invoiceeMgtKey(���޹޴��� ������ȣ) | invoiceePrintYN(���޹޴��� �μ⿩��) | closeDownState(���޹޴��� ���������) |" + _
         "closeDownStateDate(���޹޴��� ���������) | trusteeCorpName (��Ź�� ��ȣ) | trusteeCorpNum (��Ź�� ����ڹ�ȣ) | trusteeMgtKey(��Ź�� ������ȣ) | " + _
         "trusteePrintYN(��Ź�� �μ⿩��) " + vbCrLf + vbCrLf
            
    Dim info As PBTIInfo
    
    For Each info In tiSearchList.list
        tmp = tmp + info.itemKey + " | "
        tmp = tmp + info.taxType + " | "
        tmp = tmp + info.writeDate + " | "
        tmp = tmp + info.regDT + " | "
        tmp = tmp + info.issueType + " | "
        tmp = tmp + info.supplyCostTotal + " | "
        tmp = tmp + info.taxTotal + " | "
        tmp = tmp + info.purposeType + " | "
        tmp = tmp + CStr(info.lateIssueYN) + " | "
        tmp = tmp + CStr(info.openYN) + " | "
        tmp = tmp + info.openDT + " | "
        tmp = tmp + info.stateMemo + " | "
        tmp = tmp + CStr(info.stateCode) + " | "
        tmp = tmp + info.ntsConfirmNum + " | "
        tmp = tmp + info.ntsresult + " | "
        tmp = tmp + info.ntssendDT + " | "
        tmp = tmp + info.ntsresultDT + " | "
        tmp = tmp + info.ntssendErrCode + " | "
        tmp = tmp + info.modifyCode + " | "
        tmp = tmp + CStr(info.interOPYN) + " | "
        tmp = tmp + info.invoicerCorpName + " | "
        tmp = tmp + info.invoicerCorpNum + " | "
        tmp = tmp + info.invoicerMgtKey + " | "
        tmp = tmp + CStr(info.invoicerPrintYN) + " | "
        tmp = tmp + info.invoiceeCorpName + " | "
        tmp = tmp + info.invoiceeCorpNum + " | "
        tmp = tmp + info.invoiceeMgtKey + " | "
        tmp = tmp + CStr(info.invoicerPrintYN) + " | "
        tmp = tmp + CStr(info.closeDownState) + " | "
        tmp = tmp + info.closeDownStateDate + " | "
        tmp = tmp + info.trusteeCorpName + " | "
        tmp = tmp + info.trusteeCorpNum + " | "
        tmp = tmp + info.trusteeMgtKey + " | "
        tmp = tmp + CStr(info.trusteePrintYN) + vbCrLf
    Next
    
    MsgBox tmp
       
End Sub

'=========================================================================
' ���ݰ�꼭�� ���¿� ���� �����̷��� Ȯ���մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/info#GetLogs
'=========================================================================
Private Sub btnGetLogs_Click()
    Dim resultList As Collection
    Dim KeyType As MgtKeyType
    Dim tmp As String
    Dim log As PBTILog
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    
    Set resultList = TaxinvoiceService.GetLogs(txtCorpNum.Text, KeyType, txtMgtKey.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "DocLogType(�α�Ÿ��) | Log(�̷�����) | ProcType(ó������) | procCorpName(ó��ȸ���) | procContactName(ó�������) | " _
        + "ProcMemo(ó���޸�) | RegDT(����Ͻ�) | IP(������) " + vbCrLf
    
    For Each log In resultList
        tmp = tmp + CStr(log.docLogType) + " | " + log.log + " | " + log.procType + " | " + log.procCorpName + " | " + log.procContactName _
        + log.procMemo + " | " + log.regDT + " | " + log.ip + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ���ݰ�꼭�� ���õ� �ȳ� ������ ������ �մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/etc#SendEmail
'=========================================================================
Private Sub btnSendEmail_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim receiverEmail As String
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    '���Ÿ����ּ�
    receiverEmail = "test@test.com"
    
    Set Response = TaxinvoiceService.SendEmail(txtCorpNum.Text, KeyType, txtMgtKey.Text, receiverEmail)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ݰ�꼭�� ���õ� �ȳ� SMS(�ܹ�) ���ڸ� �������ϴ� �Լ���, �˺� ����Ʈ [���ڡ��ѽ�] > [����] > [���۳���] �޴����� ���۰���� Ȯ�� �� �� �ֽ��ϴ�.
' - �޽����� �ִ� 90byte���� �Է� �����ϰ�, �ʰ��� ������ �ڵ����� �����Ǿ� �����մϴ�. (�ѱ� �ִ� 45��)
' - �Լ� ȣ��� ����Ʈ�� ���ݵ˴ϴ�
' - https://developers.popbill.com/reference/taxinvoice/vb/api/etc#SendSMS
'=========================================================================
Private Sub btnSendSMS_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim sendNum As String
    Dim receiveNum As String
    Dim Contents As String
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    ' �߽Ź�ȣ
    sendNum = "07043042991"
    
    ' ���Ź�ȣ
    receiveNum = "010-111-222"
    
    ' �޽��� ����, �ִ� 90Byte (�ѱ� 45��), ���̸� �ʰ��� ������ �����Ǿ� ���۵˴ϴ�.
    Contents = "��ũ��꿡�� ���ݰ�꼭�� �����Ͽ����ϴ�. ����Ȯ�� �ٶ��ϴ�."
        
    
    Set Response = TaxinvoiceService.SendSMS(txtCorpNum.Text, KeyType, txtMgtKey.Text, _
                            sendNum, receiveNum, Contents)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ݰ�꼭�� �ѽ��� �����ϴ� �Լ���, �˺� ����Ʈ [���ڡ��ѽ�] > [�ѽ�] > [���۳���] �޴����� ���۰���� Ȯ�� �� �� �ֽ��ϴ�.
' - �Լ� ȣ��� ����Ʈ�� ���ݵ˴ϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/etc#SendFAX
'=========================================================================
Private Sub btnSendFAX_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim sendNum As String
    Dim receiveNum As String
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    '�߽��� ��ȣ
    sendNum = "07043042991"
    
    '������ �ѽ� ��ȣ
    receiveNum = "010-222-4444"
    
    Set Response = TaxinvoiceService.SendFax(txtCorpNum.Text, KeyType, txtMgtKey.Text, sendNum, receiveNum)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �˺� ����Ʈ�� ���� �����Ͽ� ������ȣ�� �ο����� ���� ���ݰ�꼭�� ������ȣ�� �Ҵ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/etc#AssignMgtKey
'=========================================================================
Private Sub btnAssignmgtkey_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim itemKey As String
    Dim mgtKey As String
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    '���ݰ�꼭 ������Ű, �����ȸ(Search) API�� ��ȯ�׸��� ItemKey ����
    itemKey = "021090515070600001"
            
    '�Ҵ��� ������ȣ, �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    mgtKey = "20220101-001"
        
    Set Response = TaxinvoiceService.AssignMgtKey(txtCorpNum.Text, KeyType, itemKey, mgtKey)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �˺� ���ڸ��� API�� ���� ������ ���ڸ����� ���ݰ�꼭�� ÷���մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/etc#AttachStatement
'=========================================================================
Private Sub btnAttachStatement_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim SubItemCode As Integer
    Dim SubMgtKey As String
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    '÷���� ���ڸ��� �����ڵ�, 121-�ŷ�����, 122-û����, 123-������, 124-���ּ�, 125-�Ա�ǥ,126-������
    SubItemCode = 121
    
    '÷���� ���ڸ��� ������ȣ
    SubMgtKey = "20220101-01"
        
    Set Response = TaxinvoiceService.AttachStatement(txtCorpNum.Text, KeyType, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ݰ�꼭�� ÷�ε� ���ڸ����� �����մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/etc#DetachStatement
'=========================================================================
Private Sub btnDetachStatement_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim SubItemCode As Integer
    Dim SubMgtKey As String
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    '÷�������� ���ڸ��� �����ڵ�, 121-�ŷ�����, 122-û����, 123-������, 124-���ּ�, 125-�Ա�ǥ, 126-������
    SubItemCode = 121
    
    '÷�������� ���ڸ��� ������ȣ
    SubMgtKey = "20220101-01"

    Set Response = TaxinvoiceService.DetachStatement(txtCorpNum.Text, KeyType, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ݰ�꼭 ���� ���� �׸� ���� �߼ۼ����� Ȯ���մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/etc#ListEmailConfig
'=========================================================================
Private Sub btnListemailconfig_Click()
    Dim resultList As Collection
    Dim i As Integer
    
    Set resultList = TaxinvoiceService.ListEmailConfig(txtCorpNum.Text)
    
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
 
    Dim tmp As String
    
    tmp = "������������(EmailType) | ���ۿ���(SendYN) " + vbCrLf
    
    Dim info As PBEmailConfig
    
    For i = 1 To resultList.Count
        If resultList(i).emailType = "TAX_ISSUE" Then
            tmp = tmp + "[������] ���޹޴��ڿ��� ���ڼ��ݰ�꼭 ���� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_ISSUE_INVOICER" Then
            tmp = tmp + "[������] �����ڿ��� ���ڼ��ݰ�꼭 ���� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_CHECK" Then
            tmp = tmp + "[������] �����ڿ��� ���ڼ��ݰ�꼭 ����Ȯ�� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_CANCEL_ISSUE" Then
            tmp = tmp + "[������] ���޹޴��ڿ��� ���ڼ��ݰ�꼭 ������� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
                    
        If resultList(i).emailType = "TAX_REQUEST" Then
            tmp = tmp + "[������] �����ڿ��� ���ݰ�꼭�� �����û �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_CANCEL_REQUEST" Then
            tmp = tmp + "[������] ���޹޴��ڿ��� ���ݰ�꼭 ��� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
            If resultList(i).emailType = "TAX_REFUSE" Then
            tmp = tmp + "[������] ���޹޴��ڿ��� ���ݰ�꼭 �ź� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_TRUST_ISSUE" Then
            tmp = tmp + "[����Ź����] ���޹޴��ڿ��� ���ڼ��ݰ�꼭 ���� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_TRUST_ISSUE_TRUSTEE" Then
            tmp = tmp + "[����Ź����] ��Ź�ڿ��� ���ڼ��ݰ�꼭 ���� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
                    
        If resultList(i).emailType = "TAX_TRUST_ISSUE_INVOICER" Then
            tmp = tmp + "[����Ź����] �����ڿ��� ���ڼ��ݰ�꼭 ���� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_TRUST_CANCEL_ISSUE" Then
            tmp = tmp + "[����Ź����] ���޹޴��ڿ��� ���ڼ��ݰ�꼭 ������� �˸� : " + resultList(i).emailType + " | "
          tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
    
        If resultList(i).emailType = "TAX_TRUST_SEND" Then
            tmp = tmp + "[����Ź����] �����ڿ��� ���ڼ��ݰ�꼭 ������� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_CLOSEDOWN" Then
            tmp = tmp + "[ó�����] �ŷ�ó�� ����� ���� Ȯ�� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_NTSFAIL_INVOICER" Then
            tmp = tmp + "[ó�����] ���ڼ��ݰ�꼭 ����û ���۽��� �ȳ�) : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
                    
        If resultList(i).emailType = "ETC_CERT_EXPIRATION" Then
            tmp = tmp + "[����߼�] �˺����� �̿����� ������������ ���� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ���ݰ�꼭 ���� ���� �׸� ���� �߼ۼ����� �����մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/etc#UpdateEmailConfig
'
' ������������
' [������]
' TAX_ISSUE : ���޹޴��ڿ��� ���ڼ��ݰ�꼭�� ���� �Ǿ����� �˷��ִ� �����Դϴ�.
' TAX_ISSUE_INVOICER : �����ڿ��� ���ڼ��ݰ�꼭�� ���� �Ǿ����� �˷��ִ� �����Դϴ�.
' TAX_CHECK : �����ڿ��� ���ڼ��ݰ�꼭�� ����Ȯ�� �Ǿ����� �˷��ִ� �����Դϴ�.
' TAX_CANCEL_ISSUE : ���޹޴��ڿ��� ���ڼ��ݰ�꼭�� ������� �Ǿ����� �˷��ִ� �����Դϴ�.
'
' [������]
' TAX_REQUEST : �����ڿ��� ���ݰ�꼭�� ���ڼ��� �Ͽ� ������ ��û�ϴ� �����Դϴ�.
' TAX_CANCEL_REQUEST : ���޹޴��ڿ��� ���ݰ�꼭�� ��� �Ǿ����� �˷��ִ� �����Դϴ�.
' TAX_REFUSE : ���޹޴��ڿ��� ���ݰ�꼭�� �ź� �Ǿ����� �˷��ִ� �����Դϴ�.
'
' [����Ź����]
' TAX_TRUST_ISSUE : ���޹޴��ڿ��� ���ڼ��ݰ�꼭�� ���� �Ǿ����� �˷��ִ� �����Դϴ�.
' TAX_TRUST_ISSUE_TRUSTEE : ��Ź�ڿ��� ���ڼ��ݰ�꼭�� ���� �Ǿ����� �˷��ִ� �����Դϴ�.
' TAX_TRUST_ISSUE_INVOICER : �����ڿ��� ���ڼ��ݰ�꼭�� ���� �Ǿ����� �˷��ִ� �����Դϴ�.
' TAX_TRUST_CANCEL_ISSUE : ���޹޴��ڿ��� ���ڼ��ݰ�꼭�� ������� �Ǿ����� �˷��ִ� �����Դϴ�.
' TAX_TRUST_CANCEL_ISSUE_INVOICER : �����ڿ��� ���ڼ��ݰ�꼭�� ������� �Ǿ����� �˷��ִ� �����Դϴ�.
'
' [ó�����]
' TAX_CLOSEDOWN : �ŷ�ó�� ����� ���θ� Ȯ���Ͽ� �ȳ��ϴ� �����Դϴ�.
' TAX_NTSFAIL_INVOICER : ���ڼ��ݰ�꼭 ����û ���۽��и� �ȳ��ϴ� �����Դϴ�.
'
' [����߼�]
' ETC_CERT_EXPIRATION : �˺����� �̿����� ������������ ������ �ȳ��ϴ� �����Դϴ�.
'
'=========================================================================
Private Sub btnUpdateemailconfig_Click()
    Dim Response As PBResponse
    Dim emailType As String
    Dim sendYN As Boolean
    
    '���� ���� ����
    emailType = "TAX_ISSUE"

    '���� ���� (True = ����, False = ������)
    sendYN = True
    
    Set Response = TaxinvoiceService.UpdateEmailConfig(txtCorpNum.Text, emailType, sendYN)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� ����û ���� �ɼ� ���� ���¸� Ȯ���մϴ�.
' - ����û ���� �ɼ� ������ �˺� ����Ʈ [���ڼ��ݰ�꼭] > [ȯ�漳��] > [���ݰ�꼭 ����] �޴����� ������ �� ������, API�� ������ �Ұ��� �մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/etc#GetSendToNTSConfig
'=========================================================================
Private Sub btnGetSendToNTSConfig_Click()
    Dim sendToNTSConfig As PBSendToNTSConfig
    
    Set sendToNTSConfig = TaxinvoiceService.GetSendToNTSConfig(txtCorpNum.Text)
    
    If sendToNTSConfig Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "����û ���� ���� : " + CStr(sendToNTSConfig.sendToNTS) + vbCrLf + "True(���� ��� ����) False(���� �ڵ� ����)"
    
End Sub

'=========================================================================
' ���ݰ�꼭 1���� �� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/view#GetPopUpURL
'=========================================================================
Private Sub btnGetPopUpURL_Click()
    Dim URL As String
    Dim KeyType As MgtKeyType
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    URL = TaxinvoiceService.GetPopUpURL(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ���ݰ�꼭 1���� �μ��ϱ� ���� �������� �˾� URL�� ��ȯ�ϸ�, ������������ �μ� �������� "������" / "���޹޴���" / "������+���޹޴���"�� �� �ϳ��� ������ �� �ֽ��ϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/view#GetPrintURL
'=========================================================================
Private Sub btnGetPrintURL_Click()
    Dim URL As String
    Dim KeyType As MgtKeyType
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    URL = TaxinvoiceService.GetPrintURL(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ���ݰ�꼭 1���� ������ ������� �μ��ϱ� ���� �������� �˾� URL�� ��ȯ�ϸ�, ������������ �μ� �������� "������" / "���޹޴���" / "������+���޹޴���"�� �� �ϳ��� ������ �� �ֽ��ϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/view#GetOldPrintURL
'=========================================================================
Private Sub btnGetOldPrintURL_Click()
Dim URL As String
    Dim KeyType As MgtKeyType
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    URL = TaxinvoiceService.GetOldPrintURL(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub
'=========================================================================
' "���޹޴���" �� ���ݰ�꼭 1���� �μ��ϱ� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/view#GetEPrintURL
'=========================================================================
Private Sub btnGetEPrintUrl_Click()
    Dim URL As String
    Dim KeyType As MgtKeyType
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    URL = TaxinvoiceService.GetEPrintURL(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' �ټ����� ���ݰ�꼭�� �μ��ϱ� ���� �������� �˾� URL�� ��ȯ�մϴ�. (�ִ� 100��)
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/view#GetMassPrintURL
'=========================================================================
Private Sub btnGetMassPrintURL_Click()
    Dim URL As String
    Dim KeyType As MgtKeyType
    Dim KeyList As New Collection
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    ' ���ڼ��ݰ�꼭 ���� ������ȣ �迭 (�ִ� 100��)
    KeyList.Add "20220101-01"
    KeyList.Add "20220101-02"
    KeyList.Add "20220101-03"
    KeyList.Add "20220101-04"
    
    URL = TaxinvoiceService.GetMassPrintURL(txtCorpNum.Text, KeyType, KeyList, txtUserID.Text)
     
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ���ڼ��ݰ�꼭 �ȳ������� �󼼺��� ��ũ URL�� ��ȯ�մϴ�.
' - �Լ� ȣ��� ��ȯ ���� URL���� ��ȿ�ð��� �����ϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/view#GetMailURL
'=========================================================================
Private Sub btnGetMailURL_Click()
    Dim URL As String
    Dim KeyType As MgtKeyType
    
    Select Case cboMgtKeyType.Text
        Case "SELL"
            KeyType = SELL
        Case "BUY"
            KeyType = BUY
        Case "TRUSTEE"
            KeyType = TRUSTEE
        Case Else
            MsgBox "������ȣ ���¸� �������ּ���."
            Exit Sub
    End Select
    
    URL = TaxinvoiceService.GetMailURL(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub
'=========================================================================
' �α��� ���·� �˺� ����Ʈ�� ���ڼ��ݰ�꼭 �ӽù����� �޴��� ������ �� �ִ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/info#GetURL
'=========================================================================
Private Sub btnGetURL_TBOX_Click(index As Integer)
    Dim URL As String
    
    URL = TaxinvoiceService.GetURL(txtCorpNum.Text, txtUserID.Text, "TBOX")
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub


'=========================================================================
' �α��� ���·� �˺� ����Ʈ�� ���ڼ��ݰ�꼭 ���� ���� ����� �޴��� ������ �� �ִ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/info#GetURL
'=========================================================================
Private Sub btnGetURL_SWBOX_Click(index As Integer)
    Dim URL As String
    
    URL = TaxinvoiceService.GetURL(txtCorpNum.Text, txtUserID.Text, "SWBOX")
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL

End Sub

'=========================================================================
' �α��� ���·� �˺� ����Ʈ�� ���ڼ��ݰ�꼭 ���� ���� ����� �޴��� ������ �� �ִ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/info#GetURL
'=========================================================================
Private Sub btnGetURL_PWBOX_Click(index As Integer)
    Dim URL As String
    
    URL = TaxinvoiceService.GetURL(txtCorpNum.Text, txtUserID.Text, "PWBOX")
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' �α��� ���·� �˺� ����Ʈ�� ���ڼ��ݰ�꼭 ���⼭�� �޴��� ������ �� �ִ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/info#GetURL
'=========================================================================
Private Sub btnGetURL_SBOX_Click()
    Dim URL As String
    
    URL = TaxinvoiceService.GetURL(txtCorpNum.Text, txtUserID.Text, "SBOX")
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
    
End Sub

'=========================================================================
' �α��� ���·� �˺� ����Ʈ�� ���ڼ��ݰ�꼭 ���Թ����� �޴��� ������ �� �ִ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/info#GetURL
'=========================================================================
Private Sub btnGetURL_PBOX_Click()
    Dim URL As String
    
    URL = TaxinvoiceService.GetURL(txtCorpNum.Text, txtUserID.Text, "PBOX")
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' �α��� ���·� �˺� ����Ʈ�� ���ڼ��ݰ�꼭 ���⹮���ۼ� �޴��� ������ �� �ִ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://developers.popbill.com/reference/taxinvoice/vb/api/info#GetURL
'=========================================================================
Private Sub btnGetURL_WRITE_Click()
    Dim URL As String
    
    URL = TaxinvoiceService.GetURL(txtCorpNum.Text, txtUserID.Text, "WRITE")
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
    
End Sub


Private Sub Form_Load()

    '��� �ʱ�ȭ
    TaxinvoiceService.Initialize LinkID, SecretKey

    '����ȯ�漳����, True-���߿� False-�����
    TaxinvoiceService.IsTest = True
    
    '������ū IP���ѱ�� ��뿩��, True-���, False-�̻��, �⺻��(True)
    TaxinvoiceService.IPRestrictOnOff = True
    
    '���ýý��� �ð� ��뿩�� True-���, False-�̻��, �⺻��(False)
    TaxinvoiceService.UseLocalTimeYN = False
    
    cboMgtKeyType.AddItem "SELL"
    cboMgtKeyType.AddItem "BUY"
    cboMgtKeyType.AddItem "TRUSTEE"
End Sub



