VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "�˺� ���ݰ�꼭 SDK ����"
   ClientHeight    =   12000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14205
   LinkTopic       =   "Form1"
   ScaleHeight     =   12000
   ScaleWidth      =   14205
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton btnPopbillURL_CERT 
      Caption         =   " ���������� ��� URL"
      Height          =   495
      Left            =   9360
      TabIndex        =   83
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton btnPopbillURL_CHRG 
      Caption         =   " ����Ʈ ���� URL"
      Height          =   495
      Left            =   9360
      TabIndex        =   82
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "���� �����ȸ"
      Height          =   390
      Left            =   3080
      TabIndex        =   81
      Top             =   10820
      Width           =   1845
   End
   Begin VB.CommandButton btnUpdateCorpInfo 
      Caption         =   "ȸ������ ����"
      Height          =   495
      Left            =   11640
      TabIndex        =   76
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton btnUpdateContact 
      Caption         =   "����� ���� ����"
      Height          =   495
      Left            =   7080
      TabIndex        =   74
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton btnListContact 
      Caption         =   "����� ��� ��ȸ"
      Height          =   495
      Left            =   7080
      TabIndex        =   73
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Frame Frame15 
      Caption         =   "ȸ������ ����"
      Height          =   1695
      Left            =   11520
      TabIndex        =   71
      Top             =   960
      Width           =   2055
      Begin VB.CommandButton btnGetCorpInfo 
         Caption         =   "ȸ������ ��ȸ"
         Height          =   495
         Left            =   120
         TabIndex        =   75
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton btnCheckID 
      Caption         =   "ID �ߺ� Ȯ��"
      Height          =   495
      Left            =   480
      TabIndex        =   69
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Frame Frame7 
      Caption         =   " ���ݰ�꼭 ���� ���"
      Height          =   8025
      Left            =   240
      TabIndex        =   14
      Top             =   3600
      Width           =   13575
      Begin VB.Frame Frame16 
         Caption         =   "������ (��ù���) ���μ���"
         Height          =   3255
         Left            =   240
         TabIndex        =   77
         Top             =   840
         Width           =   3255
         Begin VB.CommandButton btnCancelIsse_2 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�������"
            Height          =   495
            Left            =   240
            Style           =   1  '�׷���
            TabIndex        =   80
            Top             =   2110
            Width           =   975
         End
         Begin VB.CommandButton btnDelete_2 
            Caption         =   "����"
            Height          =   495
            Left            =   1920
            Style           =   1  '�׷���
            TabIndex        =   79
            Top             =   2110
            Width           =   975
         End
         Begin VB.CommandButton btnRegistIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "��ù���"
            Height          =   495
            Left            =   480
            Style           =   1  '�׷���
            TabIndex        =   78
            Top             =   720
            Width           =   1215
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
            Top             =   600
            Width           =   2625
         End
         Begin VB.Line Line18 
            X1              =   720
            X2              =   720
            Y1              =   2400
            Y2              =   960
         End
      End
      Begin VB.CommandButton btnGetEmailPublicKeys 
         Caption         =   "������ϸ��"
         Height          =   390
         Left            =   9960
         TabIndex        =   67
         Top             =   255
         Width           =   1725
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   150
         Top             =   4245
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame14 
         Caption         =   " ���� ���� "
         Height          =   2760
         Left            =   7560
         TabIndex        =   62
         Top             =   5040
         Width           =   3210
         Begin VB.CommandButton btnGetEPrintUrl 
            Caption         =   "���޹޴��� �μ� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   68
            Top             =   1260
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPopUpURL 
            Caption         =   "���� ���� ���� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   66
            Top             =   390
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPrintURL 
            Caption         =   "�μ� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   65
            Top             =   825
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMassPrintURL 
            Caption         =   "�ٷ� �μ� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   64
            Top             =   1710
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMailURL 
            Caption         =   "�̸���(���޹޴���) ��ũ URL"
            Height          =   390
            Left            =   210
            TabIndex        =   63
            Top             =   2160
            Width           =   2745
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   " ��Ÿ URL "
         Height          =   2295
         Left            =   11040
         TabIndex        =   57
         Top             =   5040
         Width           =   2265
         Begin VB.CommandButton btnGetURL_TBOX 
            Caption         =   "�ӽ� ������"
            Height          =   390
            Left            =   210
            TabIndex        =   61
            Top             =   390
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_SBOX 
            Caption         =   "���� ������"
            Height          =   390
            Left            =   210
            TabIndex        =   60
            Top             =   825
            Width           =   1845
         End
         Begin VB.CommandButton btn_GetURL_PBOX 
            Caption         =   "���� ������"
            Height          =   390
            Left            =   195
            TabIndex        =   59
            Top             =   1260
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_WRITE 
            Caption         =   "���� �ۼ�"
            Height          =   390
            Left            =   195
            TabIndex        =   58
            Top             =   1710
            Width           =   1845
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   " �ΰ� ����"
         Height          =   2775
         Left            =   5040
         TabIndex        =   53
         Top             =   5040
         Width           =   2265
         Begin VB.CommandButton btnDetachStatement 
            Caption         =   "���ڸ��� ÷������"
            Height          =   390
            Left            =   210
            TabIndex        =   85
            Top             =   2200
            Width           =   1845
         End
         Begin VB.CommandButton btnAttachStatement 
            Caption         =   "���ڸ��� ÷��"
            Height          =   390
            Left            =   210
            TabIndex        =   84
            Top             =   1750
            Width           =   1845
         End
         Begin VB.CommandButton btnSendEmail 
            Caption         =   "�̸��� ����"
            Height          =   390
            Left            =   210
            TabIndex        =   56
            Top             =   390
            Width           =   1845
         End
         Begin VB.CommandButton btnSendSMS 
            Caption         =   "���� ����"
            Height          =   390
            Left            =   210
            TabIndex        =   55
            Top             =   825
            Width           =   1845
         End
         Begin VB.CommandButton btnSendFAX 
            Caption         =   "�ѽ� ����"
            Height          =   390
            Left            =   210
            TabIndex        =   54
            Top             =   1290
            Width           =   1845
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   " ���� ���� "
         Height          =   2775
         Left            =   2640
         TabIndex        =   48
         Top             =   5040
         Width           =   2265
         Begin VB.CommandButton btnGetDetailInfo 
            Caption         =   "���� �� ����"
            Height          =   390
            Left            =   195
            TabIndex        =   52
            Top             =   1710
            Width           =   1845
         End
         Begin VB.CommandButton btnGetLogs 
            Caption         =   "���� �̷�"
            Height          =   390
            Left            =   195
            TabIndex        =   51
            Top             =   1260
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfos 
            Caption         =   "���� ����(�뷮)"
            Height          =   390
            Left            =   210
            TabIndex        =   50
            Top             =   825
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfo 
            Caption         =   "���� ����"
            Height          =   390
            Left            =   210
            TabIndex        =   49
            Top             =   390
            Width           =   1845
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   " ÷������ "
         Height          =   2280
         Left            =   240
         TabIndex        =   43
         Top             =   5055
         Width           =   2265
         Begin VB.CommandButton btnDeleteFile 
            Caption         =   "���� ����"
            Height          =   390
            Left            =   210
            TabIndex        =   47
            Top             =   1650
            Width           =   1845
         End
         Begin VB.TextBox txtFileID 
            Height          =   330
            Left            =   210
            TabIndex        =   46
            Text            =   "���Ͼ��̵�"
            Top             =   1245
            Width           =   1845
         End
         Begin VB.CommandButton btnGetFiles 
            Caption         =   "÷�� ���"
            Height          =   390
            Left            =   210
            TabIndex        =   45
            Top             =   795
            Width           =   1845
         End
         Begin VB.CommandButton btnAttachFile 
            Caption         =   "���� ÷��"
            Height          =   390
            Left            =   210
            TabIndex        =   44
            Top             =   345
            Width           =   1845
         End
      End
      Begin VB.CommandButton btnSendToNTS 
         BackColor       =   &H00C0C0FF&
         Caption         =   "����û ��� ����"
         Height          =   495
         Left            =   4680
         Style           =   1  '�׷���
         TabIndex        =   42
         Top             =   4320
         Width           =   3000
      End
      Begin VB.Frame Frame9 
         Caption         =   " ������ ���ݰ�꼭 ���μ��� "
         Height          =   3345
         Left            =   9360
         TabIndex        =   22
         Top             =   840
         Width           =   3855
         Begin VB.CommandButton btnRefuse 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�ź�"
            Height          =   375
            Left            =   2205
            Style           =   1  '�׷���
            TabIndex        =   41
            Top             =   1530
            Width           =   855
         End
         Begin VB.CommandButton btnRequestCancel 
            BackColor       =   &H00FFFFC0&
            Caption         =   "��û���"
            Height          =   375
            Left            =   2205
            Style           =   1  '�׷���
            TabIndex        =   40
            Top             =   1050
            Width           =   855
         End
         Begin VB.CommandButton btnCancelIssue_rev 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�������"
            Height          =   375
            Left            =   675
            Style           =   1  '�׷���
            TabIndex        =   39
            Top             =   2535
            Width           =   855
         End
         Begin VB.CommandButton btnDelete_rev 
            Caption         =   "����"
            Height          =   375
            Left            =   2670
            Style           =   1  '�׷���
            TabIndex        =   38
            Top             =   2550
            Width           =   855
         End
         Begin VB.CommandButton btnIssue_rev 
            BackColor       =   &H00C0C0FF&
            Caption         =   "����"
            Height          =   360
            Left            =   735
            Style           =   1  '�׷���
            TabIndex        =   37
            Top             =   1980
            Width           =   720
         End
         Begin VB.CommandButton btnRequest 
            BackColor       =   &H00FFFFC0&
            Caption         =   "��)�����û"
            Height          =   660
            Left            =   420
            Style           =   1  '�׷���
            TabIndex        =   36
            Top             =   1155
            Width           =   1350
         End
         Begin VB.CommandButton btnUpdate_rev 
            BackColor       =   &H00FFFFC0&
            Caption         =   "����"
            Height          =   375
            Left            =   2475
            Style           =   1  '�׷���
            TabIndex        =   34
            Top             =   465
            Width           =   855
         End
         Begin VB.CommandButton btnRegister_rev 
            BackColor       =   &H00FFFFC0&
            Caption         =   "���"
            Height          =   375
            Left            =   1515
            Style           =   1  '�׷���
            TabIndex        =   33
            Top             =   465
            Width           =   855
         End
         Begin VB.Line Line16 
            X1              =   1635
            X2              =   3350
            Y1              =   1740
            Y2              =   1740
         End
         Begin VB.Line Line15 
            X1              =   1650
            X2              =   3365
            Y1              =   1245
            Y2              =   1260
         End
         Begin VB.Line Line14 
            X1              =   1290
            X2              =   2925
            Y1              =   2700
            Y2              =   2700
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "�ӽ�����"
            Height          =   180
            Left            =   675
            TabIndex        =   35
            Top             =   540
            Width           =   720
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
         Begin VB.Line Line13 
            X1              =   1095
            X2              =   1095
            Y1              =   2685
            Y2              =   840
         End
         Begin VB.Line Line17 
            X1              =   3360
            X2              =   3360
            Y1              =   2730
            Y2              =   870
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "������ (�ӽ����� ����, ���࿹��) ���μ���"
         Height          =   3345
         Left            =   3720
         TabIndex        =   21
         Top             =   840
         Width           =   5415
         Begin VB.CommandButton btnCancelSend 
            BackColor       =   &H00C0C0FF&
            Caption         =   "���"
            Height          =   375
            Left            =   3930
            Style           =   1  '�׷���
            TabIndex        =   32
            Top             =   1995
            Width           =   615
         End
         Begin VB.CommandButton btnDeny 
            BackColor       =   &H00FFFFC0&
            Caption         =   "�ź�"
            Height          =   375
            Left            =   3210
            Style           =   1  '�׷���
            TabIndex        =   31
            Top             =   1995
            Width           =   615
         End
         Begin VB.CommandButton btnAccept 
            BackColor       =   &H00FFFFC0&
            Caption         =   "����"
            Height          =   375
            Left            =   2490
            Style           =   1  '�׷���
            TabIndex        =   30
            Top             =   1995
            Width           =   615
         End
         Begin VB.CommandButton btnSend 
            BackColor       =   &H00C0C0FF&
            Caption         =   "���࿹��"
            Height          =   375
            Left            =   1650
            Style           =   1  '�׷���
            TabIndex        =   29
            Top             =   1425
            Width           =   855
         End
         Begin VB.CommandButton btnRegister 
            BackColor       =   &H00C0C0FF&
            Caption         =   "���"
            Height          =   375
            Left            =   1305
            Style           =   1  '�׷���
            TabIndex        =   27
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton btnUpdate 
            BackColor       =   &H00C0C0FF&
            Caption         =   "����"
            Height          =   375
            Left            =   2265
            Style           =   1  '�׷���
            TabIndex        =   26
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton btnDelete 
            Caption         =   "����"
            Height          =   375
            Left            =   3465
            Style           =   1  '�׷���
            TabIndex        =   25
            Top             =   2760
            Width           =   855
         End
         Begin VB.CommandButton btnCancelIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�������"
            Height          =   375
            Left            =   360
            Style           =   1  '�׷���
            TabIndex        =   24
            Top             =   2730
            Width           =   855
         End
         Begin VB.CommandButton btnIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "����"
            Height          =   525
            Left            =   345
            Style           =   1  '�׷���
            TabIndex        =   23
            Top             =   1920
            Width           =   1020
         End
         Begin VB.Line Line12 
            X1              =   4080
            X2              =   4080
            Y1              =   2265
            Y2              =   2895
         End
         Begin VB.Line Line11 
            X1              =   3630
            X2              =   3630
            Y1              =   2295
            Y2              =   2925
         End
         Begin VB.Line Line10 
            X1              =   1260
            X2              =   2625
            Y1              =   2200
            Y2              =   2200
         End
         Begin VB.Line Line9 
            X1              =   4245
            X2              =   4245
            Y1              =   2100
            Y2              =   1605
         End
         Begin VB.Line Line8 
            X1              =   3480
            X2              =   3480
            Y1              =   2145
            Y2              =   1605
         End
         Begin VB.Line Line7 
            X1              =   2775
            X2              =   2775
            Y1              =   2145
            Y2              =   1605
         End
         Begin VB.Line Line4 
            X1              =   2280
            X2              =   4260
            Y1              =   1605
            Y2              =   1605
         End
         Begin VB.Line Line6 
            X1              =   1140
            X2              =   2070
            Y1              =   2115
            Y2              =   2115
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "�ӽ�����"
            Height          =   180
            Left            =   465
            TabIndex        =   28
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
            Width           =   4920
         End
         Begin VB.Line Line5 
            X1              =   2055
            X2              =   2055
            Y1              =   2100
            Y2              =   780
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            FillColor       =   &H00E0E0E0&
            Height          =   1380
            Left            =   1500
            Top             =   1245
            Width           =   3405
         End
         Begin VB.Line Line3 
            X1              =   5055
            X2              =   5055
            Y1              =   2895
            Y2              =   780
         End
         Begin VB.Line Line2 
            X1              =   900
            X2              =   5055
            Y1              =   2910
            Y2              =   2910
         End
         Begin VB.Line Line1 
            X1              =   840
            X2              =   840
            Y1              =   3000
            Y2              =   720
         End
      End
      Begin VB.ComboBox cboMgtKeyType 
         Height          =   300
         Left            =   2520
         TabIndex        =   20
         Text            =   "SELL"
         Top             =   300
         Width           =   1335
      End
      Begin VB.CommandButton btnCheckMgtKeyInUse 
         Caption         =   "������ȣ ��뿩�� Ȯ��"
         Height          =   375
         Left            =   6840
         TabIndex        =   19
         Top             =   240
         Width           =   2190
      End
      Begin VB.TextBox txtMgtKey 
         Height          =   330
         Left            =   3960
         TabIndex        =   18
         Top             =   285
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "����������ȣ( MgtKey) : "
         Height          =   180
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " �˺� �⺻ API "
      Height          =   2895
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   13575
      Begin VB.Frame Frame6 
         Caption         =   " ���������� ����"
         Height          =   1095
         Left            =   4440
         TabIndex        =   13
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnCertificateExpireDate 
            Caption         =   "������ ������ Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " ȸ������"
         Height          =   2295
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "���� ���� Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "ȸ�� ����"
            Height          =   495
            Left            =   120
            TabIndex        =   12
            Top             =   1560
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " ����Ʈ ����"
         Height          =   2295
         Left            =   1920
         TabIndex        =   8
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "��Ʈ�� �ܿ�����Ʈ Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   70
            Top             =   1560
            Width           =   2175
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ�����Ʈ Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "��� �ܰ� Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   9
            Top             =   960
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "����� ����"
         Height          =   2295
         Left            =   6720
         TabIndex        =   7
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "����� �߰�"
            Height          =   495
            Left            =   120
            TabIndex        =   72
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " �˺� �⺻ URL"
         Height          =   2295
         Left            =   9000
         TabIndex        =   5
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnGetPopbillURL_LOGIN 
            Caption         =   " �˺� �α��� URL"
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1935
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
Option Explicit

'��ũ���̵�
Private Const LinkID = "TESTER"
'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

Private TaxinvoiceService As New PBTIService

Private Sub btn_GetURL_PBOX_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetURL(txtCorpNum.Text, txtUserID.Text, "PBOX")
    
    If url = "" Then
         MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnAccept_Click()
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
    
    Set Response = TaxinvoiceService.Accept(txtCorpNum.Text, KeyType, txtMgtKey.Text, "���࿹�� ���� �޸�", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

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
    
    Set Response = TaxinvoiceService.AttachFile(txtCorpNum.Text, KeyType, txtMgtKey.Text, FilePath, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
    
End Sub

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
    
    SubItemCode = 121           '÷���� ���ڸ��� �����ڵ�, 121-�ŷ�����, 122-û����, 123-������, 124-���ּ�, 125-�Ա�ǥ,126-������
    SubMgtKey = "20151223-01"   '÷���� ���ڸ��� ������ȣ
        
    Set Response = TaxinvoiceService.AttachStatement(txtCorpNum.Text, KeyType, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnCancelIsse_2_Click()
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
    
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, "���� ��� �޸�", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnCancelIssue_Click()
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
    
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, "���� ��� �޸�", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnCancelIssue_rev_Click()
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
    
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, "���� ��� �޸�", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnCancelSend_Click()
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
    
    Set Response = TaxinvoiceService.CancelSend(txtCorpNum.Text, KeyType, txtMgtKey.Text, "���࿹�� ��� �޸�", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnCertificateExpireDate_Click()
    Dim expireDate As String
    
    expireDate = TaxinvoiceService.GetCertificateExpireDate(txtCorpNum.Text)
    
    If expireDate = "" Then
        
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������������ : " + expireDate
 
End Sub

Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.CheckIsMember(txtCorpNum.Text, LinkID)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnCheckMgtKeyInUse_Click()
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
    
    Set Response = TaxinvoiceService.CheckMgtKeyInUse(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
    
End Sub

Private Sub btnDelete_2_Click()
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
    
    Set Response = TaxinvoiceService.Delete(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

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
    
    Set Response = TaxinvoiceService.Delete(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

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
    
    Set Response = TaxinvoiceService.Delete(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

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
    
    Set Response = TaxinvoiceService.DeleteFile(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtFileID.Text, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnDeny_Click()
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
    
    Set Response = TaxinvoiceService.Deny(txtCorpNum.Text, KeyType, txtMgtKey.Text, "���࿹�� �ź� �޸�", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

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
    
    SubItemCode = 121           '÷���� ���ڸ��� �����ڵ�, 121-�ŷ�����, 122-û����, 123-������, 124-���ּ�, 125-�Ա�ǥ,126-������
    SubMgtKey = "20151223-01"   '÷���� ���ڸ��� ������ȣ
        
    Set Response = TaxinvoiceService.DetachStatement(txtCorpNum.Text, KeyType, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = TaxinvoiceService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
End Sub

Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    
    Set CorpInfo = TaxinvoiceService.GetCorpInfo(txtCorpNum.Text, txtUserID.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "ceoname : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
    
End Sub

Private Sub btnGetDetailInfo_Click()
    Dim tiDetailInfo As PBTaxinvoice
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
    
    Set tiDetailInfo = TaxinvoiceService.GetDetailInfo(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
     
    If tiDetailInfo Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "writeDate : " + tiDetailInfo.writeDate + vbCrLf
    tmp = tmp + "chargeDirection : " + tiDetailInfo.chargeDirection + vbCrLf
    tmp = tmp + "issueType : " + tiDetailInfo.issueType + vbCrLf
    tmp = tmp + "issueTiming : " + tiDetailInfo.issueTiming + vbCrLf
    tmp = tmp + "taxType : " + tiDetailInfo.taxType + vbCrLf
    
    tmp = tmp + "invoicerCorpNum : " + tiDetailInfo.invoicerCorpNum + vbCrLf
    tmp = tmp + "invoicerMgtKey : " + tiDetailInfo.invoicerMgtKey + vbCrLf
    tmp = tmp + "invoicerTaxRegID : " + tiDetailInfo.invoicerTaxRegID + vbCrLf
    tmp = tmp + "invoicerCorpName : " + tiDetailInfo.invoicerCorpName + vbCrLf
    tmp = tmp + "invoicerCEOName : " + tiDetailInfo.invoicerCEOName + vbCrLf
    tmp = tmp + "invoicerAddr : " + tiDetailInfo.invoicerAddr + vbCrLf
    tmp = tmp + "invoicerBizClass : " + tiDetailInfo.invoicerBizClass + vbCrLf
    tmp = tmp + "invoicerBizType : " + tiDetailInfo.invoicerBizType + vbCrLf
    tmp = tmp + "invoicerContactName : " + tiDetailInfo.invoicerContactName + vbCrLf
    tmp = tmp + "invoicerDeptName : " + tiDetailInfo.invoicerDeptName + vbCrLf
    tmp = tmp + "invoicerTEL : " + tiDetailInfo.invoicerTEL + vbCrLf
    tmp = tmp + "invoicerHP : " + tiDetailInfo.invoicerHP + vbCrLf
    tmp = tmp + "invoicerEmail : " + tiDetailInfo.invoicerEmail + vbCrLf
    tmp = tmp + "invoicerSMSSendYN : " + CStr(tiDetailInfo.invoicerSMSSendYN) + vbCrLf
    

    tmp = tmp + "invoiceeType : " + tiDetailInfo.invoiceeType + vbCrLf
    tmp = tmp + "invoiceeCorpNum : " + tiDetailInfo.invoiceeCorpNum + vbCrLf
    tmp = tmp + "invoiceeMgtKey : " + tiDetailInfo.invoiceeMgtKey + vbCrLf
    tmp = tmp + "invoiceeTaxRegID : " + tiDetailInfo.invoiceeTaxRegID + vbCrLf
    tmp = tmp + "invoiceeCorpName : " + tiDetailInfo.invoiceeCorpName + vbCrLf
    tmp = tmp + "invoiceeCEOName : " + tiDetailInfo.invoiceeCEOName + vbCrLf
    tmp = tmp + "invoiceeAddr : " + tiDetailInfo.invoiceeAddr + vbCrLf
    tmp = tmp + "invoiceeBizClass : " + tiDetailInfo.invoiceeBizClass + vbCrLf
    tmp = tmp + "invoiceeBizType : " + tiDetailInfo.invoiceeBizType + vbCrLf
    tmp = tmp + "invoiceeContactName1 : " + tiDetailInfo.invoiceeContactName1 + vbCrLf
    tmp = tmp + "invoiceeDeptName1 : " + tiDetailInfo.invoiceeDeptName1 + vbCrLf
    tmp = tmp + "invoiceeTEL1 : " + tiDetailInfo.invoiceeTEL1 + vbCrLf
    tmp = tmp + "invoiceeHP1 : " + tiDetailInfo.invoiceeHP1 + vbCrLf
    tmp = tmp + "invoiceeEmail1 : " + tiDetailInfo.invoiceeEmail1 + vbCrLf

    '''  �󼼳��� ���� '''
    
    MsgBox tmp
    
End Sub

Private Sub btnGetEmailPublicKeys_Click()
    Dim resultList As Collection
    
    Set resultList = TaxinvoiceService.GetEmailPublicKeys(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    Dim email As Variant
    
    For Each email In resultList
        tmp = tmp + email + vbCrLf
    Next
    
    MsgBox tmp
End Sub

Private Sub btnGetEPrintUrl_Click()
    Dim url As String
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
    
    url = TaxinvoiceService.GetEPrintURL(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetFiles_Click()
    Dim resultList As Collection
    Dim KeyType As MgtKeyType
    
    Set resultList = TaxinvoiceService.GetFiles(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "serialNum | attachedfile | displayName |  RegDT" + vbCrLf
    
    Dim file As PBAttachFile
    
    For Each file In resultList
        tmp = tmp + CStr(file.serialNum) + " | " + file.AttachedFile + " | " + file.DisplayName + " | " + file.regDT + vbCrLf
        txtFileID.Text = file.AttachedFile
    Next
    
    MsgBox tmp
End Sub

Private Sub btnGetInfo_Click()
    Dim tiInfo As PBTIInfo
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
    
    Set tiInfo = TaxinvoiceService.GetInfo(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
     
    If tiInfo Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "itemKey : " + tiInfo.itemKey + vbCrLf
    tmp = tmp + "stateCode : " + CStr(tiInfo.stateCode) + vbCrLf
    tmp = tmp + "taxType : " + tiInfo.taxType + vbCrLf
    tmp = tmp + "purposeType : " + tiInfo.purposeType + vbCrLf
    tmp = tmp + "modifyCode : " + tiInfo.modifyCode + vbCrLf
    tmp = tmp + "issueType : " + tiInfo.issueType + vbCrLf
    tmp = tmp + "lateIssueYN : " + CStr(tiInfo.lateIssueYN) + vbCrLf
    
    tmp = tmp + "writeDate : " + tiInfo.writeDate + vbCrLf
    
    tmp = tmp + "invoicerCorpName : " + tiInfo.invoicerCorpName + vbCrLf
    tmp = tmp + "invoicerCorpNum : " + tiInfo.invoicerCorpNum + vbCrLf
    tmp = tmp + "invoicerMgtKey : " + tiInfo.invoicerMgtKey + vbCrLf
    tmp = tmp + "invoiceeCorpName : " + tiInfo.invoiceeCorpName + vbCrLf
    tmp = tmp + "invoiceeCorpNum : " + tiInfo.invoiceeCorpNum + vbCrLf
    tmp = tmp + "invoiceeMgtKey : " + tiInfo.invoiceeMgtKey + vbCrLf
    tmp = tmp + "trusteeCorpName : " + tiInfo.trusteeCorpName + vbCrLf
    tmp = tmp + "trusteeCorpNum : " + tiInfo.trusteeCorpNum + vbCrLf
    tmp = tmp + "trusteeMgtKey : " + tiInfo.trusteeMgtKey + vbCrLf
    
    tmp = tmp + "supplyCostTotal : " + tiInfo.supplyCostTotal + vbCrLf
    tmp = tmp + "taxTotal : " + tiInfo.taxTotal + vbCrLf
    
    tmp = tmp + "issueDT : " + tiInfo.issueDT + vbCrLf
    tmp = tmp + "preIssueDT : " + tiInfo.preIssueDT + vbCrLf
    tmp = tmp + "stateDT : " + tiInfo.stateDT + vbCrLf
    tmp = tmp + "openYN : " + CStr(tiInfo.openYN) + vbCrLf
    tmp = tmp + "openDT : " + tiInfo.openDT + vbCrLf
    
    tmp = tmp + "ntsresult : " + tiInfo.ntsresult + vbCrLf
    tmp = tmp + "ntsconfirmNum : " + tiInfo.ntsconfirmNum + vbCrLf
    tmp = tmp + "ntssendDT : " + tiInfo.ntssendDT + vbCrLf
    tmp = tmp + "ntsresultDT : " + tiInfo.ntsresultDT + vbCrLf
    tmp = tmp + "ntssendErrCode : " + tiInfo.ntssendErrCode + vbCrLf
    
    tmp = tmp + "stateMemo : " + tiInfo.stateMemo + vbCrLf
    
    tmp = tmp + "regDT : " + tiInfo.regDT + vbCrLf
    
    
    MsgBox tmp
    
    
End Sub

Private Sub btnGetInfos_Click()
    Dim resultList As Collection
    Dim KeyType As MgtKeyType
    Dim KeyList As New Collection
    
    KeyType = SELL
    
    '������ȣ �迭, �ִ� 1000��
    KeyList.Add "123123"
    KeyList.Add "123123"
    KeyList.Add "123"
    KeyList.Add "123123123"
    
    Set resultList = TaxinvoiceService.GetInfos(txtCorpNum.Text, KeyType, KeyList, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "ItemKey | StateCode | TaxType | WriteDate | RegDT" + vbCrLf
    
    Dim info As PBTIInfo
    
    For Each info In resultList
        tmp = tmp + info.itemKey + " | " + CStr(info.stateCode) + " | " + info.taxType + " | " + info.writeDate + " | " + info.regDT + vbCrLf
    Next
    
    MsgBox tmp
End Sub

Private Sub btnGetLogs_Click()
    Dim resultList As Collection
    Dim KeyType As MgtKeyType
    
    Set resultList = TaxinvoiceService.GetLogs(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "DocLogType | Log | ProcType | ProcCorpName | ProcMemo | RegDT | IP" + vbCrLf
    
    Dim log As PBTILog
    
    For Each log In resultList
        tmp = tmp + CStr(log.docLogType) + " | " + log.log + " | " + log.procType + " | " + log.procCorpName + " | " + log.procMemo + " | " + log.regDT + " | " + log.ip + vbCrLf
    Next
    
    MsgBox tmp
End Sub

Private Sub btnGetMailURL_Click()
    Dim url As String
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
    
    url = TaxinvoiceService.GetMailURL(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetMassPrintURL_Click()
    Dim url As String
    Dim KeyType As MgtKeyType
    Dim KeyList As New Collection
    
    KeyType = SELL
    KeyList.Add "123123"
    KeyList.Add "123123"
    KeyList.Add "123"
    KeyList.Add "123123123"
    
    url = TaxinvoiceService.GetMassPrintURL(txtCorpNum.Text, KeyType, KeyList, txtUserID.Text)
     
    If url = "" Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    
End Sub

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = TaxinvoiceService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
    
End Sub
Private Sub btnGetPopbillURL_LOGIN_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
         MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetPopUpURL_Click()
    Dim url As String
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
    
    url = TaxinvoiceService.GetPopUpURL(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
    
End Sub

Private Sub btnGetPrintURL_Click()
    Dim url As String
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
    
    url = TaxinvoiceService.GetPrintURL(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetURL_SBOX_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetURL(txtCorpNum.Text, txtUserID.Text, "SBOX")
    
    If url = "" Then
         MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetURL_TBOX_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetURL(txtCorpNum.Text, txtUserID.Text, "TBOX")
    
    If url = "" Then
         MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetURL_WRITE_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetURL(txtCorpNum.Text, txtUserID.Text, "WRITE")
    
    If url = "" Then
         MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnIssue_Click()
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
    
    Set Response = TaxinvoiceService.Issue(txtCorpNum.Text, KeyType, txtMgtKey.Text, "����޸�", "", True, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
    
End Sub

Private Sub btnIssue_rev_Click()
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
    
    Set Response = TaxinvoiceService.Issue(txtCorpNum.Text, KeyType, txtMgtKey.Text, "����޸�", "", False, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    joinData.LinkID = LinkID '��ũ ���̵�
    joinData.CorpNum = "1231212312" '����ڹ�ȣ "-" ����.
    joinData.ceoname = "��ǥ�ڼ���"
    joinData.corpName = "ȸ����ȣ"
    joinData.addr = "�ּ�"
    joinData.bizType = "����"
    joinData.bizClass = "����"
    joinData.id = "userid"      '6�� �̻� 20�� �̸�.
    joinData.pwd = "pwd_must_be_long_enough"    '6�� �̻� 20�� �̸�.
    joinData.ContactName = "����ڼ���"
    joinData.ContactTEL = "02-999-9999"
    joinData.ContactHP = "010-1234-5678"
    joinData.ContactFAX = "02-999-9998"
    joinData.ContactEmail = "test@test.com"
    
    Set Response = TaxinvoiceService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
    
    
End Sub

Private Sub btnListContact_Click()
    Dim resultList As Collection
        
    Set resultList = TaxinvoiceService.ListContact(txtCorpNum.Text, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "id | email | hp | personName | searchAllAllowYN | tel | fax | mgrYN | regDT " + vbCrLf
    
    Dim info As PBContactInfo
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.email + " | " + info.hp + " | " + info.personName + " | " + CStr(info.searchAllAllowYN) _
                + info.tel + " | " + info.fax + " | " + CStr(info.mgrYN) + " | " + info.regDT + vbCrLf
    Next
    
    MsgBox tmp
End Sub

Private Sub btnPopbillURL_CERT_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CERT")
    
    If url = "" Then
         MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnPopbillURL_CHRG_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CHRG")
    
    If url = "" Then
         MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnRefuse_Click()
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
    
    Set Response = TaxinvoiceService.Refuse(txtCorpNum.Text, KeyType, txtMgtKey.Text, "��)���� ��û �ź� �޸�", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    joinData.id = "testkorea_20151007"      '����� ���̵�
    joinData.pwd = "test@test.com"          '��й�ȣ
    joinData.personName = "����ڸ�"        '����ڸ�
    joinData.tel = "070-1234-1234"          '����ó
    joinData.hp = "010-1234-1234"           '�޴�����ȣ
    joinData.email = "test@test.com"        '�̸��� �ּ�
    joinData.fax = "070-1234-1234"          '�ѽ���ȣ
    joinData.searchAllAllowYN = True        '��ü��ȸ����, Ture-ȸ����ȸ, False-������ȸ
    joinData.mgrYN = False                  '������ ���ѿ���
        
    Set Response = TaxinvoiceService.RegistContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
    
End Sub

Private Sub btnRegister_Click()
    Dim Taxinvoice As New PBTaxinvoice
    
    Taxinvoice.writeDate = "20151012"             '�ʼ�, ����� �ۼ�����
    Taxinvoice.chargeDirection = "������"         '�ʼ�, {������, ������}
    Taxinvoice.issueType = "������"               '�ʼ�, {������, ������, ����Ź}
    Taxinvoice.purposeType = "����"               '�ʼ�, {����, û��}
    Taxinvoice.issueTiming = "��������"           '�ʼ�, {��������, ���ν��ڵ�����}
    Taxinvoice.taxType = "����"                   '�ʼ�, {����, ����, �鼼}
        
    Taxinvoice.invoicerCorpNum = "1234567890"     '������ ����ڹ�ȣ
    Taxinvoice.invoicerTaxRegID = ""              '������� �ĺ���ȣ. �ʿ�� ����. ������ ���� 4�ڸ�.
    Taxinvoice.invoicerCorpName = "������ ��ȣ"
    Taxinvoice.invoicerMgtKey = txtMgtKey.Text    '������ ����������ȣ, 1~24�ڸ�, ����,����,'-','_' �����Ͽ� ���Ƿ� ����
    Taxinvoice.invoicerCEOName = "������"" ��ǥ�� ����"
    Taxinvoice.invoicerAddr = "������ �ּ�"
    Taxinvoice.invoicerBizClass = "������ ����"
    Taxinvoice.invoicerBizType = "������ ����,����2"
    Taxinvoice.invoicerContactName = "������ ����ڸ�"
    Taxinvoice.invoicerEmail = "test@test.com"
    Taxinvoice.invoicerTEL = "070-7070-0707"
    Taxinvoice.invoicerHP = "010-000-2222"
    Taxinvoice.invoicerSMSSendYN = False            '�������(������->���޹޴���) ���ڹ߼ۿ���
    
    Taxinvoice.invoiceeType = "�����"             '���޹޴��� ����, {�����, ����, �ܱ���} �� ����
    Taxinvoice.invoiceeCorpNum = "8888888888"      '���޹޴��� ����ڹ�ȣ
    Taxinvoice.invoiceeCorpName = "���޹޴��� ��ȣ"
    Taxinvoice.invoiceeMgtKey = ""                 '���޹޴��� ����������ȣ(������ÿ��� �ʼ�)
    Taxinvoice.invoiceeCEOName = "���޹޴��� ��ǥ�� ����"
    Taxinvoice.invoiceeAddr = "���޹޴��� �ּ�"
    Taxinvoice.invoiceeBizClass = "���޹޴��� ����"
    Taxinvoice.invoiceeBizType = "���޹޴��� ����"
    Taxinvoice.invoiceeContactName1 = "���޹޴��� ����ڸ�"
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    Taxinvoice.invoiceeHP1 = "010-111-222"
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    Taxinvoice.invoiceeSMSSendYN = False          '�������(���޹޴���->������) ���ڹ߼ۿ���
            
    Taxinvoice.supplyCostTotal = "100000"         '�ʼ� ���ް��� �հ�
    Taxinvoice.taxTotal = "10000"                 '�ʼ� ���� �հ�
    Taxinvoice.totalAmount = "110000"             '�ʼ� �հ�ݾ�.  ���ް��� + ����
    
    Taxinvoice.modifyCode = "" '�������ݰ�꼭 �ۼ��� 1~6���� ���ñ���.
    Taxinvoice.originalTaxinvoiceKey = "" '�������ݰ�꼭 �ۼ��� �������ݰ�꼭�� ItemKey����. ItemKey�� ����Ȯ��.
    Taxinvoice.serialNum = "123"  '�Ϸù�ȣ
    Taxinvoice.cash = ""          '����
    Taxinvoice.chkBill = ""       '��ǥ
    Taxinvoice.note = ""          '����
    Taxinvoice.credit = ""        '�ܻ�̼���
    Taxinvoice.remark1 = "���1"
    Taxinvoice.remark2 = "���2"
    Taxinvoice.remark3 = "���3"
    Taxinvoice.kwon = "1"           '��
    Taxinvoice.ho = "1"             'ȣ
    
    Taxinvoice.businessLicenseYN = False '����ڵ���� �̹��� ÷�ν� ����.
    Taxinvoice.bankBookYN = False         '����纻 �̹��� ÷�ν� ����.
    
    
    '���׸� �߰�.
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20140410"   '�ŷ�����  yyyyMMdd
    newDetail.itemName = "ǰ��"
    newDetail.spec = "�԰�"
    newDetail.qty = "1" '����           ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.unitCost = "100000"       ' �ܰ�, �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.supplyCost = "100000"     ' ���ް���
    newDetail.tax = "10000"             ' ����
    newDetail.remark = "���"           ' ���
    
    Taxinvoice.detailList.Add newDetail
    
    Set newDetail = New PBTIDetail
    newDetail.serialNum = 2
    newDetail.itemName = "ǰ��2"
    
    Taxinvoice.detailList.Add newDetail
    
    
    '�߰������ �߰�. �ɼ�.
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                '�Ϸù�ȣ, 1���� ��������
    newContact.ContactName = "����� ����"
    newContact.email = "test2@test.com"
    
    Taxinvoice.addContactList.Add newContact
        
    
    Dim Response As PBResponse
    
    'Register(����ڹ�ȣ, ���ݰ�꼭 ��ü, �ŷ����� �����ۼ�����, �˺�ȸ�����̵�)
    Set Response = TaxinvoiceService.Register(txtCorpNum.Text, Taxinvoice, False, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
    

End Sub

Private Sub btnRegister_rev_Click()
    Dim Taxinvoice As New PBTaxinvoice
    
    Taxinvoice.writeDate = "20151008"             '�ʼ�, ����� �ۼ�����
    Taxinvoice.chargeDirection = "������"         '�ʼ�, {������, ������}
    Taxinvoice.issueType = "������"               '�ʼ�, {������, ������, ����Ź}
    Taxinvoice.purposeType = "����"               '�ʼ�, {����, û��}
    Taxinvoice.issueTiming = "��������"           '�ʼ�, {��������, ���ν��ڵ�����}
    Taxinvoice.taxType = "����"                   '�ʼ�, {����, ����, �鼼}
    
    
    Taxinvoice.invoicerCorpNum = "8888888888"
    Taxinvoice.invoicerTaxRegID = "" '������� �ĺ���ȣ. �ʿ�� ����. ������ ���� 4�ڸ�.
    Taxinvoice.invoicerCorpName = "������ ��ȣ"
    Taxinvoice.invoicerMgtKey = ""
    Taxinvoice.invoicerCEOName = "������"" ��ǥ�� ����"
    Taxinvoice.invoicerAddr = "������ �ּ�"
    Taxinvoice.invoicerBizClass = "������ ����"
    Taxinvoice.invoicerBizType = "������ ����,����2"
    Taxinvoice.invoicerContactName = "������ ����ڸ�"
    Taxinvoice.invoicerEmail = "test@test.com"
    Taxinvoice.invoicerTEL = "070-7070-0707"
    Taxinvoice.invoicerHP = "010-000-2222"
    Taxinvoice.invoicerSMSSendYN = True '����� ���ڹ߼۱�� ���� Ȱ��
    
    Taxinvoice.invoiceeType = "�����"
    Taxinvoice.invoiceeCorpNum = "1231212312"
    Taxinvoice.invoiceeCorpName = "���޹޴��� ��ȣ"
    Taxinvoice.invoiceeMgtKey = txtMgtKey.Text
    Taxinvoice.invoiceeCEOName = "���޹޴��� ��ǥ�� ����"
    Taxinvoice.invoiceeAddr = "���޹޴��� �ּ�"
    Taxinvoice.invoiceeBizClass = "���޹޴��� ����"
    Taxinvoice.invoiceeBizType = "���޹޴��� ����"
    Taxinvoice.invoiceeContactName1 = "���޹޴��� ����ڸ�"
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    Taxinvoice.supplyCostTotal = "100000"         '�ʼ� ���ް��� �հ�
    Taxinvoice.taxTotal = "10000"                 '�ʼ� ���� �հ�
    Taxinvoice.totalAmount = "110000"             '�ʼ� �հ�ݾ�.  ���ް��� + ����
    
    Taxinvoice.modifyCode = "" '�������ݰ�꼭 �ۼ��� 1~6���� ���ñ���.
    Taxinvoice.originalTaxinvoiceKey = "" '�������ݰ�꼭 �ۼ��� �������ݰ�꼭�� ItemKey����. ItemKey�� ����Ȯ��.
    Taxinvoice.serialNum = "123"
    Taxinvoice.cash = ""          '����
    Taxinvoice.chkBill = ""       '��ǥ
    Taxinvoice.note = ""          '����
    Taxinvoice.credit = ""        '�ܻ�̼���
    Taxinvoice.remark1 = "���1"
    Taxinvoice.remark2 = "���2"
    Taxinvoice.remark3 = "���3"
    Taxinvoice.kwon = "1"
    Taxinvoice.ho = "1"
    
    Taxinvoice.businessLicenseYN = False '����ڵ���� �̹��� ÷�ν� ����.
    Taxinvoice.bankBookYN = False         '����纻 �̹��� ÷�ν� ����.
       
    
    '���׸� �߰�.
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1
    newDetail.purchaseDT = "20140410"
    newDetail.itemName = "ǰ��"
    newDetail.spec = "�԰�"
    newDetail.qty = "1" '����
    newDetail.unitCost = "100000"
    newDetail.supplyCost = "100000"
    newDetail.tax = "10000"
    newDetail.remark = "���"
    
    Taxinvoice.detailList.Add newDetail
    
    Set newDetail = New PBTIDetail
    newDetail.serialNum = 2
    newDetail.itemName = "ǰ��2"
    
    Taxinvoice.detailList.Add newDetail
    
    
    '�߰������ �߰�. �ɼ�.
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.ContactName = "����� ����"
    newContact.email = "test2@test.com"
    
    Taxinvoice.addContactList.Add newContact
    
    
    Dim Response As PBResponse
    'Register(����ڹ�ȣ, ���ݰ�꼭 ��ü, �ŷ����� �����ۼ�����, �˺�ȸ�����̵�)
    Set Response = TaxinvoiceService.Register(txtCorpNum.Text, Taxinvoice, False, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
End Sub

Private Sub btnRegistIssue_Click()
    
    Dim Taxinvoice As New PBTaxinvoice
    
    Taxinvoice.writeSpecification = False         '�ŷ����� �����ۼ� ����
    Taxinvoice.forceIssue = False                 '�������� ��������
    Taxinvoice.memo = ""                          '�޸�
    Taxinvoice.emailSubject = ""                  '�ȳ����� ����, ����ó���� �⺻�������� ����
    Taxinvoice.dealInvoiceMgtKey = ""             '�ŷ����� �����ۼ��� ���� ������ȣ, �̱���� ���ݰ�꼭 ������ȣ�� �ڵ��ۼ�
        
    Taxinvoice.writeDate = "20151012"             '�ʼ�, ����� �ۼ�����
    Taxinvoice.chargeDirection = "������"         '�ʼ�, {������, ������}
    Taxinvoice.issueType = "������"               '�ʼ�, {������, ������, ����Ź}
    Taxinvoice.purposeType = "����"               '�ʼ�, {����, û��}
    Taxinvoice.issueTiming = "��������"           '�ʼ�, {��������, ���ν��ڵ�����}
    Taxinvoice.taxType = "����"                   '�ʼ�, {����, ����, �鼼}
        
    Taxinvoice.invoicerCorpNum = "1234567890"     '������ ����ڹ�ȣ
    Taxinvoice.invoicerTaxRegID = ""              '������� �ĺ���ȣ. �ʿ�� ����. ������ ���� 4�ڸ�.
    Taxinvoice.invoicerCorpName = "������ ��ȣ"
    Taxinvoice.invoicerMgtKey = txtMgtKey.Text    '������ ����������ȣ, 1~24�ڸ�, ����,����,'-','_' �����Ͽ� ���Ƿ� ����
    Taxinvoice.invoicerCEOName = "������"" ��ǥ�� ����"
    Taxinvoice.invoicerAddr = "������ �ּ�"
    Taxinvoice.invoicerBizClass = "������ ����"
    Taxinvoice.invoicerBizType = "������ ����,����2"
    Taxinvoice.invoicerContactName = "������ ����ڸ�"
    Taxinvoice.invoicerEmail = "test@test.com"
    Taxinvoice.invoicerTEL = "070-7070-0707"
    Taxinvoice.invoicerHP = "010-000-2222"
    Taxinvoice.invoicerSMSSendYN = False            '�������(������->���޹޴���) ���ڹ߼ۿ���
    
    Taxinvoice.invoiceeType = "�����"             '���޹޴��� ����, {�����, ����, �ܱ���} �� ����
    Taxinvoice.invoiceeCorpNum = "8888888888"      '���޹޴��� ����ڹ�ȣ
    Taxinvoice.invoiceeCorpName = "���޹޴��� ��ȣ"
    Taxinvoice.invoiceeMgtKey = ""                 '���޹޴��� ����������ȣ(������ÿ��� �ʼ�)
    Taxinvoice.invoiceeCEOName = "���޹޴��� ��ǥ�� ����"
    Taxinvoice.invoiceeAddr = "���޹޴��� �ּ�"
    Taxinvoice.invoiceeBizClass = "���޹޴��� ����"
    Taxinvoice.invoiceeBizType = "���޹޴��� ����"
    Taxinvoice.invoiceeContactName1 = "���޹޴��� ����ڸ�"
    Taxinvoice.invoiceeEmail1 = "test@test.com"
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    Taxinvoice.invoiceeHP1 = "010-111-222"
    Taxinvoice.invoiceeSMSSendYN = False          '�������(���޹޴���->������) ���ڹ߼ۿ���
            
    Taxinvoice.supplyCostTotal = "100000"         '�ʼ� ���ް��� �հ�
    Taxinvoice.taxTotal = "10000"                 '�ʼ� ���� �հ�
    Taxinvoice.totalAmount = "110000"             '�ʼ� �հ�ݾ�.  ���ް��� + ����
    
    Taxinvoice.modifyCode = "" '�������ݰ�꼭 �ۼ��� 1~6���� ���ñ���.
    Taxinvoice.originalTaxinvoiceKey = "" '�������ݰ�꼭 �ۼ��� �������ݰ�꼭�� ItemKey����. ItemKey�� ����Ȯ��.
    Taxinvoice.serialNum = "123"  '�Ϸù�ȣ
    Taxinvoice.cash = ""          '����
    Taxinvoice.chkBill = ""       '��ǥ
    Taxinvoice.note = ""          '����
    Taxinvoice.credit = ""        '�ܻ�̼���
    Taxinvoice.remark1 = "���1"
    Taxinvoice.remark2 = "���2"
    Taxinvoice.remark3 = "���3"
    Taxinvoice.kwon = "1"           '��
    Taxinvoice.ho = "1"             'ȣ
    
    Taxinvoice.businessLicenseYN = False '����ڵ���� �̹��� ÷�ν� ����.
    Taxinvoice.bankBookYN = False         '����纻 �̹��� ÷�ν� ����.
    
    
    '���׸� �߰�.
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20140410"   '�ŷ�����  yyyyMMdd
    newDetail.itemName = "ǰ��"
    newDetail.spec = "�԰�"
    newDetail.qty = "1" '����           ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.unitCost = "100000"       '�ܰ� �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.supplyCost = "100000"     '���ް���
    newDetail.tax = "10000"             '����
    newDetail.remark = "���"           '���
    
    Taxinvoice.detailList.Add newDetail
    
    Set newDetail = New PBTIDetail
    newDetail.serialNum = 2
    newDetail.itemName = "ǰ��2"
    
    Taxinvoice.detailList.Add newDetail
    
    
    '�߰������ �߰�. �ɼ�.
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                '�Ϸù�ȣ 1���� ��������
    newContact.ContactName = "����� ����"  '����ڸ�
    newContact.email = "test2@test.com"     '�̸����ּ�
    
    Taxinvoice.addContactList.Add newContact
        
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.RegistIssue(txtCorpNum.Text, Taxinvoice, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
    
End Sub

Private Sub btnRequest_Click()
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
    
    Set Response = TaxinvoiceService.Request(txtCorpNum.Text, KeyType, txtMgtKey.Text, "��)���� ��û �޸�", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnRequestCancel_Click()
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
    
    Set Response = TaxinvoiceService.CancelRequest(txtCorpNum.Text, KeyType, txtMgtKey.Text, "��)���� ��û ��� �޸�", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnSearch_Click()
    Dim tiSearchList As PBTISearchList
    Dim KeyType As MgtKeyType
    Dim DType As String
    Dim SDate As String
    Dim EDate As String
    Dim State As New Collection
    Dim TType As New Collection
    Dim taxType As New Collection
    Dim LateOnly As String
    Dim Page As Integer
    Dim PerPage As Integer
    Dim Order As String
    
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
    
    
    DType = "I"             '[�ʼ�] ��������, R-����Ͻ� W-�ۼ����� I-�����Ͻ� �� ��1
    SDate = "20151206"      '[�ʼ�] ��������, yyyyMMdd
    EDate = "20151231"      '[�ʼ�] ��������, yyyyMMdd
    
    State.Add "100"         '���ۻ��°� �迭, �̱���� ��ü������ȸ, �������°� 3�ڸ����� �ۼ�
    State.Add "2**"         '2,3��° ���ϵ�ī�� ����
    State.Add "3**"
    
    TType.Add "N"           '�������� �迭, N-�Ϲ� M-���� �� ����, �̱���� ��ü��ȸ
    TType.Add "M"
    
    taxType.Add "T"         '�������� �迭, T-����, N-�鼼 Z-���� �� ����, �̱���� ��ü��ȸ
    taxType.Add "N"
    taxType.Add "Z"
    
    LateOnly = ""           '�������� ����, 0-�������и� ��ȸ 1-��������и���ȸ, ����ó���� ��ü��ȸ
    
    Page = 1                '������ ��ȣ
    PerPage = 10            '������ ��ϰ���, �ִ� 1000��
    
    Order = "A"             '���Ĺ���, D-��������(�⺻��), A-��������
    
    Set tiSearchList = TaxinvoiceService.Search(txtCorpNum.Text, KeyType, DType, SDate, EDate, State, TType, taxType, LateOnly, Page, PerPage, Order)
     
    If tiSearchList Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    tmp = "code : " + CStr(tiSearchList.code) + vbCrLf
    tmp = tmp + "total : " + CStr(tiSearchList.total) + vbCrLf
    tmp = tmp + "perPage : " + CStr(tiSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum : " + CStr(tiSearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount : " + CStr(tiSearchList.pageCount) + vbCrLf
    tmp = tmp + "message : " + tiSearchList.message + vbCrLf + vbCrLf
    
    tmp = tmp + "itemKey | stateCode | TaxTye | writeDate | regDT | lateIssueYN | invoicerCorpNum | invoicerCorpName | invoiceeCorpNum | invoiceeCorpName | " + _
                "issueType | supplyCostTotal | taxTotal" + vbCrLf
            
    Dim info As PBTIInfo
    
    For Each info In tiSearchList.list
        tmp = tmp + info.itemKey + " | " + CStr(info.stateCode) + " | " + info.taxType + " | " + info.writeDate + " | " + info.regDT + " | " + _
        CStr(info.lateIssueYN) + " | " + info.invoicerCorpNum + " | " + info.invoicerCorpName + " | " + info.invoiceeCorpNum + " | " + info.invoiceeCorpName + " | " + _
        info.issueType + " | " + info.supplyCostTotal + " | " + info.taxTotal + vbCrLf
    Next
    
    MsgBox tmp
       
End Sub

Private Sub btnSend_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim emailSubject As String
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
    
    emailSubject = ""       '���࿹�� �ȳ����� ����, ����ó���� �⺻�������� ����
    memo = "���࿹�� �޸�"  '�޸�
    
    Set Response = TaxinvoiceService.Send(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, emailSubject, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnSendEmail_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim receiveEmail As String
    
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
    
    receiveEmail = "test@test.com" '������ �����ּ�
    
    Set Response = TaxinvoiceService.SendEmail(txtCorpNum.Text, KeyType, txtMgtKey.Text, receiveEmail, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnSendFAX_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim senderNum As String
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
        
    senderNum = "07075103710"     '�߽Ź�ȣ
    receiveNum = "111-222-4444"   '���Ź�ȣ
        
    Set Response = TaxinvoiceService.SendFax(txtCorpNum.Text, KeyType, txtMgtKey.Text, senderNum, receiveNum, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnSendSMS_Click()
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim senderNum As String
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
    
    senderNum = "07075103710"
    receiveNum = "111-2222-4444"
    Contents = "���� ����, 90Byte�ʰ��� ���̰� �����Ǿ� ���۵�"
    
    Set Response = TaxinvoiceService.SendSMS(txtCorpNum.Text, KeyType, txtMgtKey.Text, senderNum, receiveNum, Contents, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

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
    
    Set Response = TaxinvoiceService.SendToNTS(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnUnitCost_Click()
    Dim unitCost As Double
    
    unitCost = TaxinvoiceService.GetUnitCost(txtCorpNum.Text)
    
    If unitCost < 0 Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "����ܰ� : " + CStr(unitCost)
End Sub

Private Sub btnUpdate_Click()

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
    
    Dim Taxinvoice As New PBTaxinvoice
    
    Taxinvoice.writeDate = "20140319"             '�ʼ�, ����� �ۼ�����
    Taxinvoice.chargeDirection = "������"         '�ʼ�, {������, ������}
    Taxinvoice.issueType = "������"               '�ʼ�, {������, ������, ����Ź}
    Taxinvoice.purposeType = "����"               '�ʼ�, {����, û��}
    Taxinvoice.issueTiming = "��������"           '�ʼ�, {��������, ���ν��ڵ�����}
    Taxinvoice.taxType = "����"                   '�ʼ�, {����, ����, �鼼}
    
    
    Taxinvoice.invoicerCorpNum = "1231212312"
    Taxinvoice.invoicerTaxRegID = "" '������� �ĺ���ȣ. �ʿ�� ����. ������ ���� 4�ڸ�.
    Taxinvoice.invoicerCorpName = "������ ��ȣ"
    Taxinvoice.invoicerMgtKey = txtMgtKey.Text
    Taxinvoice.invoicerCEOName = "������"" ��ǥ�� ����"
    Taxinvoice.invoicerAddr = "������ �ּ�"
    Taxinvoice.invoicerBizClass = "������ ����"
    Taxinvoice.invoicerBizType = "������ ����,����2"
    Taxinvoice.invoicerContactName = "������ ����ڸ�"
    Taxinvoice.invoicerEmail = "test@test.com"
    Taxinvoice.invoicerTEL = "070-7070-0707"
    Taxinvoice.invoicerHP = "010-000-2222"
    Taxinvoice.invoicerSMSSendYN = True '����� ���ڹ߼۱�� ���� Ȱ��
    
    Taxinvoice.invoiceeType = "�����"
    Taxinvoice.invoiceeCorpNum = "8888888888"
    Taxinvoice.invoiceeCorpName = "���޹޴��� ��ȣ"
    Taxinvoice.invoiceeMgtKey = ""
    Taxinvoice.invoiceeCEOName = "���޹޴��� ��ǥ�� ����"
    Taxinvoice.invoiceeAddr = "���޹޴��� �ּ�"
    Taxinvoice.invoiceeBizClass = "���޹޴��� ����"
    Taxinvoice.invoiceeBizType = "���޹޴��� ����"
    Taxinvoice.invoiceeContactName1 = "���޹޴��� ����ڸ�"
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    Taxinvoice.supplyCostTotal = "100000"         '�ʼ� ���ް��� �հ�
    Taxinvoice.taxTotal = "10000"                 '�ʼ� ���� �հ�
    Taxinvoice.totalAmount = "110000"             '�ʼ� �հ�ݾ�.  ���ް��� + ����
    
    Taxinvoice.modifyCode = "" '�������ݰ�꼭 �ۼ��� 1~6���� ���ñ���.
    Taxinvoice.originalTaxinvoiceKey = "" '�������ݰ�꼭 �ۼ��� �������ݰ�꼭�� ItemKey����. ItemKey�� ����Ȯ��.
    Taxinvoice.serialNum = "123"
    Taxinvoice.cash = ""          '����
    Taxinvoice.chkBill = ""       '��ǥ
    Taxinvoice.note = ""          '����
    Taxinvoice.credit = ""        '�ܻ�̼���
    Taxinvoice.remark1 = "���1"
    Taxinvoice.remark2 = "���2"
    Taxinvoice.remark3 = "���3"
    Taxinvoice.kwon = "1"
    Taxinvoice.ho = "1"
    
    Taxinvoice.businessLicenseYN = False '����ڵ���� �̹��� ÷�ν� ����.
    Taxinvoice.bankBookYN = False         '����纻 �̹��� ÷�ν� ����.
        
    
    '���׸� �߰�.
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1
    newDetail.purchaseDT = "20140410"
    newDetail.itemName = "ǰ��"
    newDetail.spec = "�԰�"
    newDetail.qty = "1" '����
    newDetail.unitCost = "100000"
    newDetail.supplyCost = "100000"
    newDetail.tax = "10000"
    newDetail.remark = "���"
    
    Taxinvoice.detailList.Add newDetail
    
    Set newDetail = New PBTIDetail
    newDetail.serialNum = 2
    newDetail.itemName = "ǰ��2_������"
    
    Taxinvoice.detailList.Add newDetail
    
    
    '�߰������ �߰�. �ɼ�.
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.ContactName = "����� ����"
    newContact.email = "test2@test.com"
    
    Taxinvoice.addContactList.Add newContact
    
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.Update(txtCorpNum.Text, KeyType, txtMgtKey.Text, Taxinvoice, False, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
End Sub


Private Function ByteArrayToHex(ByRef ByteArray() As Byte) As String
    Dim l As Long, strRet As String
    
    For l = LBound(ByteArray) To UBound(ByteArray)
        strRet = strRet & Hex$(ByteArray(l)) & " "
    Next l
    
    'Remove last space at end.
    ByteArrayToHex = Left$(strRet, Len(strRet) - 1)
End Function

Private Sub btnUpdate_rev_Click()

    Dim KeyType As MgtKeyType
    
    KeyType = BUY
    
    Dim Taxinvoice As New PBTaxinvoice
    
    Taxinvoice.writeDate = "20140319"             '�ʼ�, ����� �ۼ�����
    Taxinvoice.chargeDirection = "������"         '�ʼ�, {������, ������}
    Taxinvoice.issueType = "������"               '�ʼ�, {������, ������, ����Ź}
    Taxinvoice.purposeType = "����"               '�ʼ�, {����, û��}
    Taxinvoice.issueTiming = "��������"           '�ʼ�, {��������, ���ν��ڵ�����}
    Taxinvoice.taxType = "����"                   '�ʼ�, {����, ����, �鼼}
    
    
    Taxinvoice.invoicerCorpNum = "8888888888"
    Taxinvoice.invoicerTaxRegID = "" '������� �ĺ���ȣ. �ʿ�� ����. ������ ���� 4�ڸ�.
    Taxinvoice.invoicerCorpName = "������ ��ȣ"
    Taxinvoice.invoicerMgtKey = ""
    Taxinvoice.invoicerCEOName = "������"" ��ǥ�� ����"
    Taxinvoice.invoicerAddr = "������ �ּ�"
    Taxinvoice.invoicerBizClass = "������ ����"
    Taxinvoice.invoicerBizType = "������ ����,����2"
    Taxinvoice.invoicerContactName = "������ ����ڸ�"
    Taxinvoice.invoicerEmail = "test@test.com"
    Taxinvoice.invoicerTEL = "070-7070-0707"
    Taxinvoice.invoicerHP = "010-000-2222"
    Taxinvoice.invoicerSMSSendYN = True '����� ���ڹ߼۱�� ���� Ȱ��
    
    Taxinvoice.invoiceeType = "�����"
    Taxinvoice.invoiceeCorpNum = "1231212312"
    Taxinvoice.invoiceeCorpName = "���޹޴��� ��ȣ"
    Taxinvoice.invoiceeMgtKey = txtMgtKey.Text
    Taxinvoice.invoiceeCEOName = "���޹޴��� ��ǥ�� ����"
    Taxinvoice.invoiceeAddr = "���޹޴��� �ּ�"
    Taxinvoice.invoiceeBizClass = "���޹޴��� ����"
    Taxinvoice.invoiceeBizType = "���޹޴��� ����"
    Taxinvoice.invoiceeContactName1 = "���޹޴��� ����ڸ�"
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    Taxinvoice.supplyCostTotal = "100000"         '�ʼ� ���ް��� �հ�
    Taxinvoice.taxTotal = "10000"                 '�ʼ� ���� �հ�
    Taxinvoice.totalAmount = "110000"             '�ʼ� �հ�ݾ�.  ���ް��� + ����
    
    Taxinvoice.modifyCode = "" '�������ݰ�꼭 �ۼ��� 1~6���� ���ñ���.
    Taxinvoice.originalTaxinvoiceKey = "" '�������ݰ�꼭 �ۼ��� �������ݰ�꼭�� ItemKey����. ItemKey�� ����Ȯ��.
    Taxinvoice.serialNum = "123"
    Taxinvoice.cash = ""          '����
    Taxinvoice.chkBill = ""       '��ǥ
    Taxinvoice.note = ""          '����
    Taxinvoice.credit = ""        '�ܻ�̼���
    Taxinvoice.remark1 = "���1"
    Taxinvoice.remark2 = "���2"
    Taxinvoice.remark3 = "���3"
    Taxinvoice.kwon = "1"
    Taxinvoice.ho = "1"
    
    Taxinvoice.businessLicenseYN = False '����ڵ���� �̹��� ÷�ν� ����.
    Taxinvoice.bankBookYN = False         '����纻 �̹��� ÷�ν� ����.
    
    '���׸� �߰�.
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1
    newDetail.purchaseDT = "20140410"
    newDetail.itemName = "ǰ��"
    newDetail.spec = "�԰�"
    newDetail.qty = "1" '����
    newDetail.unitCost = "100000"
    newDetail.supplyCost = "100000"
    newDetail.tax = "10000"
    newDetail.remark = "���"
    
    Taxinvoice.detailList.Add newDetail
    
    Set newDetail = New PBTIDetail
    newDetail.serialNum = 2
    newDetail.itemName = "ǰ��2_������"
    
    Taxinvoice.detailList.Add newDetail
    
    
    '�߰������ �߰�. �ɼ�.
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.ContactName = "����� ����"
    newContact.email = "test2@test.com"
    
    Taxinvoice.addContactList.Add newContact
    
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.Update(txtCorpNum.Text, KeyType, txtMgtKey.Text, Taxinvoice, False, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
End Sub

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    joinData.personName = "����ڸ�_����"  '����ڸ�
    joinData.tel = "070-1234-1234"         '����ó
    joinData.hp = "010-1234-1234"          '�޴�����ȣ
    joinData.email = "test@test.com"       '�̸��� �ּ�
    joinData.fax = "070-1234-1234"         '�ѽ���ȣ
    joinData.searchAllAllowYN = True       '��ü��ȸ����, Ture-ȸ����ȸ, False-������
    joinData.mgrYN = False                 '������ ���ѿ���
                
    Set Response = TaxinvoiceService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    CorpInfo.ceoname = "��ǥ��"         '��ǥ�ڸ�
    CorpInfo.corpName = "��ȣ"          '��ȣ��
    CorpInfo.addr = "����Ư����"        '�ּ�
    CorpInfo.bizType = "����"           '����
    CorpInfo.bizClass = "����"          '����
    
    Set Response = TaxinvoiceService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub


Private Sub Form_Load()
    '��� �ʱ�ȭ
    TaxinvoiceService.Initialize LinkID, SecretKey
    
    '����ȯ�� ������ True(�׽�Ʈ��), False(�����)
    TaxinvoiceService.IsTest = True
        
    cboMgtKeyType.AddItem "SELL"
    cboMgtKeyType.AddItem "BUY"
    cboMgtKeyType.AddItem "TRUSTEE"
End Sub

