VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "�˺� ���ݰ�꼭 SDK ����"
   ClientHeight    =   11865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19065
   LinkTopic       =   "Form1"
   ScaleHeight     =   11865
   ScaleWidth      =   19065
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton btnUpdateemailconfig 
      Caption         =   "�˸����� ���ۼ��� ����"
      Height          =   390
      Index           =   3
      Left            =   7560
      TabIndex        =   91
      Top             =   9960
      Width           =   2085
   End
   Begin VB.CommandButton btnListemailconfig 
      Caption         =   "�˸����� ���۸�� ��ȸ"
      Height          =   390
      Index           =   2
      Left            =   7560
      TabIndex        =   90
      Top             =   9480
      Width           =   2085
   End
   Begin VB.CommandButton btnAssignmgtkey 
      Caption         =   "������ȣ �Ҵ�"
      Height          =   390
      Index           =   1
      Left            =   5520
      TabIndex        =   89
      Top             =   9960
      Width           =   1965
   End
   Begin VB.Frame Frame17 
      Caption         =   "��Ʈ�ʰ��� ����Ʈ"
      Height          =   1935
      Index           =   1
      Left            =   6720
      TabIndex        =   84
      Top             =   840
      Width           =   2295
      Begin VB.CommandButton btnGetPartnerURL_CHRG 
         Caption         =   "����Ʈ ���� URL"
         Height          =   410
         Left            =   120
         TabIndex        =   88
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton btnGetPartnerBalance 
         Caption         =   "��Ʈ�� �ܿ�����Ʈ Ȯ��"
         Height          =   410
         Left            =   120
         TabIndex        =   87
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "���� �����ȸ"
      Height          =   390
      Left            =   3080
      TabIndex        =   77
      Top             =   10800
      Width           =   1845
   End
   Begin VB.CommandButton btnUpdateCorpInfo 
      Caption         =   "ȸ������ ����"
      Height          =   410
      Left            =   9360
      TabIndex        =   72
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton btnUpdateContact 
      Caption         =   "����� ���� ����"
      Height          =   410
      Left            =   13800
      TabIndex        =   70
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton btnListContact 
      Caption         =   "����� ��� ��ȸ"
      Height          =   410
      Left            =   13800
      TabIndex        =   69
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Frame Frame15 
      Caption         =   "ȸ������ ����"
      Height          =   1935
      Left            =   9240
      TabIndex        =   67
      Top             =   840
      Width           =   2055
      Begin VB.CommandButton btnGetCorpInfo 
         Caption         =   "ȸ������ ��ȸ"
         Height          =   410
         Left            =   120
         TabIndex        =   71
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton btnCheckID 
      Caption         =   "ID �ߺ� Ȯ��"
      Height          =   410
      Left            =   480
      TabIndex        =   66
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   " �˺� �⺻ API "
      Height          =   2295
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   18495
      Begin VB.Frame Frame17 
         Caption         =   "�������� ����Ʈ"
         Height          =   1935
         Index           =   0
         Left            =   4440
         TabIndex        =   83
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   86
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   " ����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   85
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " ���������� ����"
         Height          =   1935
         Left            =   11160
         TabIndex        =   12
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton btnCheckCertValidation 
            Caption         =   "������ ��ȿ�� Ȯ��"
            Height          =   375
            Left            =   120
            TabIndex        =   94
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton btnGetTaxCertURL 
            Caption         =   " ������ ��� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   81
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
         Height          =   1935
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
         Height          =   1935
         Left            =   1920
         TabIndex        =   8
         Top             =   240
         Width           =   2415
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "�������� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   80
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
         Height          =   1935
         Left            =   13440
         TabIndex        =   7
         Top             =   240
         Width           =   2055
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "����� �߰�"
            Height          =   410
            Left            =   120
            TabIndex        =   68
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " �˺� �⺻ URL"
         Height          =   1935
         Left            =   15600
         TabIndex        =   5
         Top             =   240
         Width           =   2655
         Begin VB.CommandButton btnGetSealURL 
            Caption         =   "�ΰ� �� ÷�ι��� ��� URL"
            Height          =   375
            Left            =   120
            TabIndex        =   82
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
   Begin VB.Frame Frame18 
      Caption         =   " (����) ��ÿ�û ���μ���"
      Height          =   3255
      Left            =   10320
      TabIndex        =   95
      Top             =   4200
      Width           =   3615
      Begin VB.CommandButton btnRegistRequest 
         BackColor       =   &H00FFFFC0&
         Caption         =   "��ÿ�û"
         Height          =   420
         Left            =   1560
         Style           =   1  '�׷���
         TabIndex        =   101
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton btnIssue_rev_sub 
         BackColor       =   &H00C0C0FF&
         Caption         =   "����"
         Height          =   420
         Left            =   330
         Style           =   1  '�׷���
         TabIndex        =   100
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton btnDelete_rev_sub 
         Caption         =   "����"
         Height          =   420
         Left            =   2520
         Style           =   1  '�׷���
         TabIndex        =   99
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton btnCancelIssue_rev_sub 
         BackColor       =   &H00C0C0FF&
         Caption         =   "�������"
         Height          =   420
         Left            =   330
         Style           =   1  '�׷���
         TabIndex        =   98
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton btnRequestCancel_sub 
         BackColor       =   &H00FFFFC0&
         Caption         =   "��û���"
         Height          =   420
         Left            =   2520
         Style           =   1  '�׷���
         TabIndex        =   97
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton btnRefuse_sub 
         BackColor       =   &H00C0C0FF&
         Caption         =   "�ź�"
         Height          =   420
         Left            =   1440
         Style           =   1  '�׷���
         TabIndex        =   96
         Top             =   1560
         Width           =   855
      End
      Begin VB.Line Line24 
         X1              =   720
         X2              =   1900
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line22 
         X1              =   1905
         X2              =   1905
         Y1              =   1200
         Y2              =   2760
      End
      Begin VB.Line Line20 
         X1              =   1200
         X2              =   2760
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line15 
         X1              =   720
         X2              =   720
         Y1              =   1200
         Y2              =   2760
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��ÿ�û"
         Height          =   180
         Left            =   480
         TabIndex        =   103
         Top             =   600
         Width           =   720
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
      Begin VB.Line Line23 
         X1              =   1320
         X2              =   1320
         Y1              =   960
         Y2              =   1200
      End
      Begin VB.Line Line21 
         X1              =   2880
         X2              =   2880
         Y1              =   960
         Y2              =   2760
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   " ���ݰ�꼭 ���� ���"
      Height          =   8625
      Left            =   240
      TabIndex        =   13
      Top             =   3000
      Width           =   18495
      Begin VB.Frame Frame16 
         Caption         =   " (����) ��ù��� ���μ���"
         Height          =   3255
         Left            =   480
         TabIndex        =   73
         Top             =   1200
         Width           =   3255
         Begin VB.CommandButton btnCancelIsse_2 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�������"
            Height          =   495
            Left            =   360
            Style           =   1  '�׷���
            TabIndex        =   76
            Top             =   2110
            Width           =   975
         End
         Begin VB.CommandButton btnDelete_2 
            Caption         =   "����"
            Height          =   495
            Left            =   1920
            Style           =   1  '�׷���
            TabIndex        =   75
            Top             =   2110
            Width           =   975
         End
         Begin VB.CommandButton btnRegistIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "��ù���"
            Height          =   495
            Left            =   360
            Style           =   1  '�׷���
            TabIndex        =   74
            Top             =   720
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
            Top             =   600
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
         Caption         =   "������ϸ��"
         Height          =   390
         Left            =   16200
         TabIndex        =   64
         Top             =   240
         Width           =   1725
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   240
         Top             =   4560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame14 
         Caption         =   " ���� ���� "
         Height          =   2760
         Left            =   9720
         TabIndex        =   59
         Top             =   5640
         Width           =   3210
         Begin VB.CommandButton btnGetEPrintUrl 
            Caption         =   "���޹޴��� �μ� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   65
            Top             =   1260
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPopUpURL 
            Caption         =   "���� ���� ���� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   63
            Top             =   390
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPrintURL 
            Caption         =   "�μ� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   62
            Top             =   825
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMassPrintURL 
            Caption         =   "�ٷ� �μ� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   61
            Top             =   1710
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMailURL 
            Caption         =   "�̸���(���޹޴���) ��ũ URL"
            Height          =   390
            Left            =   210
            TabIndex        =   60
            Top             =   2160
            Width           =   2745
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   " ��Ÿ URL "
         Height          =   2295
         Left            =   13200
         TabIndex        =   54
         Top             =   5640
         Width           =   2265
         Begin VB.CommandButton btnGetURL_TBOX 
            Caption         =   "�ӽ� ������"
            Height          =   390
            Left            =   210
            TabIndex        =   58
            Top             =   390
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_SBOX 
            Caption         =   "���� ������"
            Height          =   390
            Left            =   210
            TabIndex        =   57
            Top             =   825
            Width           =   1845
         End
         Begin VB.CommandButton btn_GetURL_PBOX 
            Caption         =   "���� ������"
            Height          =   390
            Left            =   195
            TabIndex        =   56
            Top             =   1260
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_WRITE 
            Caption         =   "���� �ۼ�"
            Height          =   390
            Left            =   195
            TabIndex        =   55
            Top             =   1710
            Width           =   1845
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   " �ΰ� ����"
         Height          =   2775
         Left            =   5040
         TabIndex        =   52
         Top             =   5640
         Width           =   4545
         Begin VB.CommandButton btnSendFAX 
            Caption         =   "�ѽ� ����"
            Height          =   375
            Left            =   240
            TabIndex        =   93
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton btnSendSMS 
            Caption         =   "���� ����"
            Height          =   375
            Left            =   240
            TabIndex        =   92
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton btnDetachStatement 
            Caption         =   "���ڸ��� ÷������"
            Height          =   390
            Left            =   2280
            TabIndex        =   79
            Top             =   840
            Width           =   2085
         End
         Begin VB.CommandButton btnAttachStatement 
            Caption         =   "���ڸ��� ÷��"
            Height          =   390
            Left            =   2280
            TabIndex        =   78
            Top             =   390
            Width           =   2085
         End
         Begin VB.CommandButton btnSendEmail 
            Caption         =   "�̸��� ����"
            Height          =   390
            Left            =   240
            TabIndex        =   53
            Top             =   390
            Width           =   1965
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   " ���� ���� "
         Height          =   2775
         Left            =   2640
         TabIndex        =   47
         Top             =   5640
         Width           =   2265
         Begin VB.CommandButton btnGetDetailInfo 
            Caption         =   "���� �� ����"
            Height          =   390
            Left            =   195
            TabIndex        =   51
            Top             =   1710
            Width           =   1845
         End
         Begin VB.CommandButton btnGetLogs 
            Caption         =   "���� �̷�"
            Height          =   390
            Left            =   195
            TabIndex        =   50
            Top             =   1260
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfos 
            Caption         =   "���� ����(�뷮)"
            Height          =   390
            Left            =   210
            TabIndex        =   49
            Top             =   825
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfo 
            Caption         =   "���� ����"
            Height          =   390
            Left            =   210
            TabIndex        =   48
            Top             =   390
            Width           =   1845
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   " ÷������ "
         Height          =   2280
         Left            =   240
         TabIndex        =   42
         Top             =   5655
         Width           =   2265
         Begin VB.CommandButton btnDeleteFile 
            Caption         =   "���� ����"
            Height          =   390
            Left            =   210
            TabIndex        =   46
            Top             =   1650
            Width           =   1845
         End
         Begin VB.TextBox txtFileID 
            Height          =   330
            Left            =   210
            TabIndex        =   45
            Text            =   "���Ͼ��̵�"
            Top             =   1245
            Width           =   1845
         End
         Begin VB.CommandButton btnGetFiles 
            Caption         =   "÷�� ���"
            Height          =   390
            Left            =   210
            TabIndex        =   44
            Top             =   795
            Width           =   1845
         End
         Begin VB.CommandButton btnAttachFile 
            Caption         =   "���� ÷��"
            Height          =   390
            Left            =   210
            TabIndex        =   43
            Top             =   345
            Width           =   1845
         End
      End
      Begin VB.CommandButton btnSendToNTS 
         BackColor       =   &H00C0C0FF&
         Caption         =   "����û ��� ����"
         Height          =   495
         Left            =   2160
         Style           =   1  '�׷���
         TabIndex        =   41
         Top             =   4560
         Width           =   4440
      End
      Begin VB.Frame Frame9 
         Caption         =   " �ӽ����� ������ ���μ��� "
         Height          =   3255
         Left            =   13920
         TabIndex        =   21
         Top             =   1200
         Width           =   4095
         Begin VB.CommandButton btnRefuse 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�ź�"
            Height          =   420
            Left            =   1320
            Style           =   1  '�׷���
            TabIndex        =   40
            Top             =   1800
            Width           =   855
         End
         Begin VB.CommandButton btnRequestCancel 
            BackColor       =   &H00FFFFC0&
            Caption         =   "��û���"
            Height          =   420
            Left            =   2760
            Style           =   1  '�׷���
            TabIndex        =   39
            Top             =   1200
            Width           =   855
         End
         Begin VB.CommandButton btnCancelIssue_rev 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�������"
            Height          =   420
            Left            =   360
            Style           =   1  '�׷���
            TabIndex        =   38
            Top             =   2520
            Width           =   855
         End
         Begin VB.CommandButton btnDelete_rev 
            Caption         =   "����"
            Height          =   420
            Left            =   2760
            Style           =   1  '�׷���
            TabIndex        =   37
            Top             =   2520
            Width           =   855
         End
         Begin VB.CommandButton btnIssue_rev 
            BackColor       =   &H00C0C0FF&
            Caption         =   "����"
            Height          =   420
            Left            =   360
            Style           =   1  '�׷���
            TabIndex        =   36
            Top             =   1800
            Width           =   855
         End
         Begin VB.CommandButton btnRequest 
            BackColor       =   &H00FFFFC0&
            Caption         =   "��)�����û"
            Height          =   420
            Left            =   320
            Style           =   1  '�׷���
            TabIndex        =   35
            Top             =   1200
            Width           =   1920
         End
         Begin VB.CommandButton btnUpdate_rev 
            BackColor       =   &H00FFFFC0&
            Caption         =   "����"
            Height          =   375
            Left            =   2475
            Style           =   1  '�׷���
            TabIndex        =   33
            Top             =   465
            Width           =   855
         End
         Begin VB.CommandButton btnRegister_rev 
            BackColor       =   &H00FFFFC0&
            Caption         =   "���"
            Height          =   375
            Left            =   1515
            Style           =   1  '�׷���
            TabIndex        =   32
            Top             =   465
            Width           =   855
         End
         Begin VB.Line Line25 
            X1              =   2040
            X2              =   2880
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line16 
            X1              =   1680
            X2              =   1680
            Y1              =   1560
            Y2              =   2760
         End
         Begin VB.Line Line14 
            X1              =   1080
            X2              =   2925
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "�ӽ�����"
            Height          =   180
            Left            =   675
            TabIndex        =   34
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
            X1              =   750
            X2              =   750
            Y1              =   2685
            Y2              =   840
         End
         Begin VB.Line Line17 
            X1              =   3240
            X2              =   3240
            Y1              =   2630
            Y2              =   1500
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   " �ӽ����� ����, ���࿹�� ���μ���"
         Height          =   3255
         Left            =   3960
         TabIndex        =   20
         Top             =   1200
         Width           =   5415
         Begin VB.CommandButton btnCancelSend 
            BackColor       =   &H00C0C0FF&
            Caption         =   "���"
            Height          =   375
            Left            =   3930
            Style           =   1  '�׷���
            TabIndex        =   31
            Top             =   1995
            Width           =   615
         End
         Begin VB.CommandButton btnDeny 
            BackColor       =   &H00FFFFC0&
            Caption         =   "�ź�"
            Height          =   375
            Left            =   3210
            Style           =   1  '�׷���
            TabIndex        =   30
            Top             =   1995
            Width           =   615
         End
         Begin VB.CommandButton btnAccept 
            BackColor       =   &H00FFFFC0&
            Caption         =   "����"
            Height          =   375
            Left            =   2490
            Style           =   1  '�׷���
            TabIndex        =   29
            Top             =   1995
            Width           =   615
         End
         Begin VB.CommandButton btnSend 
            BackColor       =   &H00C0C0FF&
            Caption         =   "���࿹��"
            Height          =   375
            Left            =   1650
            Style           =   1  '�׷���
            TabIndex        =   28
            Top             =   1425
            Width           =   855
         End
         Begin VB.CommandButton btnRegister 
            BackColor       =   &H00C0C0FF&
            Caption         =   "���"
            Height          =   375
            Left            =   1305
            Style           =   1  '�׷���
            TabIndex        =   26
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton btnUpdate 
            BackColor       =   &H00C0C0FF&
            Caption         =   "����"
            Height          =   375
            Left            =   2265
            Style           =   1  '�׷���
            TabIndex        =   25
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton btnDelete 
            Caption         =   "����"
            Height          =   375
            Left            =   3345
            Style           =   1  '�׷���
            TabIndex        =   24
            Top             =   2760
            Width           =   975
         End
         Begin VB.CommandButton btnCancelIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�������"
            Height          =   375
            Left            =   360
            Style           =   1  '�׷���
            TabIndex        =   23
            Top             =   2730
            Width           =   975
         End
         Begin VB.CommandButton btnIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "����"
            Height          =   495
            Left            =   360
            Style           =   1  '�׷���
            TabIndex        =   22
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "���࿹��"
            Height          =   180
            Left            =   4080
            TabIndex        =   106
            Top             =   1320
            Width           =   720
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
            X2              =   4250
            Y1              =   1605
            Y2              =   1605
         End
         Begin VB.Line Line6 
            X1              =   1090
            X2              =   2050
            Y1              =   2115
            Y2              =   2115
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "�ӽ�����"
            Height          =   180
            Left            =   465
            TabIndex        =   27
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
            Y2              =   790
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
         TabIndex        =   19
         Text            =   "SELL"
         Top             =   300
         Width           =   1335
      End
      Begin VB.CommandButton btnCheckMgtKeyInUse 
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
      Begin VB.Frame Frame19 
         Caption         =   " ������ ���μ���"
         Height          =   4575
         Left            =   240
         TabIndex        =   102
         Top             =   720
         Width           =   9375
      End
      Begin VB.Frame Frame20 
         Caption         =   " ������ ���μ���"
         Height          =   4575
         Left            =   9840
         TabIndex        =   104
         Top             =   720
         Width           =   8415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "����������ȣ( MgtKey) : "
         Height          =   180
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   2115
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�ӽ�����"
      Height          =   180
      Left            =   8040
      TabIndex        =   105
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
' - VB6 SDK ����ȯ�� ������� �ȳ� : http://blog.linkhub.co.kr/569
' - ������Ʈ ���� : 2018-11-21
' - ���� ������� ����ó : 1600-9854 / 070-4304-2991
' - ���� ������� �̸��� : code@linkhub.co.kr
'
' <�׽�Ʈ �������� �غ����>
' 1) 30, 33�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
' 2) �˺� ���߿� ����Ʈ(test.popbill.com)�� ����ȸ������ �����մϴ�.
' 3) ���ڼ��ݰ�꼭 ������ ���� ������������ ����մϴ�.
'    - �˺�����Ʈ �α��� > [���ڼ��ݰ�꼭] > [ȯ�漳��]
'      > [���������� ����]
'    - ���������� ��� �˾� URL (GetTaxCertURL API)�� �̿��Ͽ� ���
'
'=========================================================================

Option Explicit

'=========================================================================
' - ��������(��ũ���̵�, ���Ű)�� ��Ʈ���� ����ȸ���� �ĺ��ϴ�
'   ������ ���Ǵ� ������ ������� �ʵ��� �����Ͻñ� �ٶ��ϴ�.
' - ����� ��ȯ���Ŀ��� ��������(��ũ���̵�, ���Ű)�� ������� �ʽ��ϴ�.
'=========================================================================

'��ũ���̵�
Private Const LinkID = "TESTER"

'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'���ݰ�꼭 ��ü ����
Private TaxinvoiceService As New PBTIService

'=========================================================================
' �˺� > ���� ������ �˾� URL�� ��ȯ�մϴ�.
' - ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
'=========================================================================

Private Sub btn_GetURL_PBOX_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetURL(txtCorpNum.Text, txtUserID.Text, "PBOX")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ���࿹�� ���ݰ�꼭�� [����]ó���մϴ�.
'=========================================================================

Private Sub btnAccept_Click()
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
    memo = "���࿹�� ���θ޸�"
    
    Set Response = TaxinvoiceService.Accept(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �˺� ����Ʈ���� �ۼ��� ���ݰ�꼭�� ��Ʈ�� ����������ȣ�� �Ҵ��մϴ�.
' - ����������ȣ�� �������� �ʴ� ���ݰ�꼭�� �Ҵ��� ���� �մϴ�.
'=========================================================================

Private Sub btnAssignmgtkey_Click(index As Integer)
    Dim Response As PBResponse
    Dim KeyType As MgtKeyType
    Dim itemKey As String
    Dim MgtKey As String
    
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
    itemKey = "018090515070600001"
            
    '�Ҵ��� ����������ȣ, ����, ����, '-', '_' ��������
    '1~24�ڸ����� ����ڹ�ȣ�� �ߺ����� ������ȣ �Ҵ�
    MgtKey = "20180911-001"
        
    Set Response = TaxinvoiceService.AssignMgtKey(txtCorpNum.Text, KeyType, itemKey, MgtKey)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ݰ�꼭�� ÷�������� ����մϴ�.
' - [�ӽ�����] ������ ���ݰ�꼭�� ������ ÷���Ҽ� �ֽ��ϴ�.
' - ÷�������� �ִ� 5������ ����� �� �ֽ��ϴ�.
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
'1���� ���ڸ����� ���ݰ�꼭�� ÷���մϴ�.
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
    SubMgtKey = "20151223-01"
        
    Set Response = TaxinvoiceService.AttachStatement(txtCorpNum.Text, KeyType, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
'[����Ϸ�] ������ ���ݰ�꼭�� [�������] ó���մϴ�.
' - [�������]�� ����û ���������� �����մϴ�.
' - ������ҵ� ���ݰ�꼭�� ����û�� ���۵��� �ʽ��ϴ�.
' - ������� ���ݰ�꼭�� ����� ����������ȣ�� ���� �ϱ� ���ؼ���
'   ����(Delete API)�� ȣ���Ͽ� [����] ó�� �ϼž� �մϴ�.
'=========================================================================

Private Sub btnCancelIsse_2_Click()
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
'[����Ϸ�] ������ ���ݰ�꼭�� [�������] ó���մϴ�.
' - [�������]�� ����û ���������� �����մϴ�.
' - ������ҵ� ���ݰ�꼭�� ����û�� ���۵��� �ʽ��ϴ�.
' - ������� ���ݰ�꼭�� ����� ����������ȣ�� ���� �ϱ� ���ؼ���
'   ����(Delete API)�� ȣ���Ͽ� [����] ó�� �ϼž� �մϴ�.
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
    memo = "���� ��� �޸�"
    
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
'[����Ϸ�] ������ ���ݰ�꼭�� [�������] ó���մϴ�.
' - [�������]�� ����û ���������� �����մϴ�.
' - ������ҵ� ���ݰ�꼭�� ����û�� ���۵��� �ʽ��ϴ�.
' - ������� ���ݰ�꼭�� ����� ����������ȣ�� ���� �ϱ� ���ؼ���
'   ����(Delete API)�� ȣ���Ͽ� [����] ó�� �ϼž� �մϴ�.
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
    memo = "���� ��� �޸�"
    
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub
'=========================================================================
'[����Ϸ�] ������ ���ݰ�꼭�� [�������] ó���մϴ�.
' - [�������]�� ����û ���������� �����մϴ�.
' - ������ҵ� ���ݰ�꼭�� ����û�� ���۵��� �ʽ��ϴ�.
' - ������� ���ݰ�꼭�� ����� ����������ȣ�� ���� �ϱ� ���ؼ���
'   ����(Delete API)�� ȣ���Ͽ� [����] ó�� �ϼž� �մϴ�.
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
    memo = "���� ��� �޸�"
    
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���࿹�� ���ݰ�꼭�� [���] ó�� �մϴ�.
' - [���]�� ���ݰ�꼭�� ����(Delete API)�ϸ� ��ϵ� ����������ȣ��
'   ������ �� �ֽ��ϴ�.
'=========================================================================

Private Sub btnCancelSend_Click()
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
    memo = "���࿹�� ��� �޸�"
    
    Set Response = TaxinvoiceService.CancelSend(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �˺��� ��ϵǾ� �ִ� ������������ �������ڸ� Ȯ���մϴ�.
' - ������������ ����/��߱�/��й�ȣ ������ �Ǵ� ��� �ش� ��������
'   ���� �ϼž� ���������� API�� �̿��Ͻ� �� �ֽ��ϴ�.
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
' �˺��� ��ϵ� ������������ ��ȿ���� Ȯ���Ѵ�.
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
' �˺� ȸ�����̵� �ߺ����θ� Ȯ���մϴ�.
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
' �ش� ������� ��Ʈ�� ����ȸ�� ���Կ��θ� Ȯ���մϴ�.
' - LinkID�� ���������� �����Ǿ� �ִ� ��ũ���̵� ���Դϴ�.
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
' ���ݰ�꼭 ������ȣ �ߺ����θ� Ȯ���մϴ�.
' - ������ȣ�� 1~24�ڸ��� ����, ���� '-', '_' �������� ������ �� �ֽ��ϴ�.
'=========================================================================

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
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
    
End Sub

'=========================================================================
' 1���� ���ڼ��ݰ�꼭�� �����մϴ�.
' - ���ݰ�꼭�� �����ؾ߸� ����������ȣ(mgtKey)�� ������ �� �ֽ��ϴ�.
' - ���������� ���� ���� : [�ӽ�����], [�������], [���࿹�� ���],
'   [���࿹�� �ź�]
'=========================================================================

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
    
    Set Response = TaxinvoiceService.Delete(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' 1���� ���ڼ��ݰ�꼭�� �����մϴ�.
' - ���ݰ�꼭�� �����ؾ߸� ����������ȣ(mgtKey)�� ������ �� �ֽ��ϴ�.
' - ���������� ���� ���� : [�ӽ�����], [�������], [���࿹�� ���],
'   [���࿹�� �ź�]
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
' 1���� ���ڼ��ݰ�꼭�� �����մϴ�.
' - ���ݰ�꼭�� �����ؾ߸� ����������ȣ(mgtKey)�� ������ �� �ֽ��ϴ�.
' - ���������� ���� ���� : [�ӽ�����], [�������], [���࿹�� ���],
'   [���࿹�� �ź�]
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
' 1���� ���ڼ��ݰ�꼭�� �����մϴ�.
' - ���ݰ�꼭�� �����ؾ߸� ����������ȣ(mgtKey)�� ������ �� �ֽ��ϴ�.
' - ���������� ���� ���� : [�ӽ�����], [�������], [���࿹�� ���],
'   [���࿹�� �ź�]
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
' ���ݰ�꼭�� ÷�ε� ������ �����մϴ�.
' - ������ �ĺ��ϴ� ���Ͼ��̵�� ÷������ ���(GetFileList API) �� �����׸�
'   �� ���Ͼ��̵�(AttachedFile) ���� ���� Ȯ���� �� �ֽ��ϴ�.
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
' ���࿹�� ���ݰ�꼭�� [�ź�]ó�� �մϴ�.
' - [�ź�]ó���� ���ݰ�꼭�� ����(Delete API)�ϸ� ��ϵ� ����������ȣ��
'   ������ �� �ֽ��ϴ�.
'=========================================================================

Private Sub btnDeny_Click()
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
    memo = "���࿹�� �ź� �޸�"
    
    Set Response = TaxinvoiceService.Deny(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
'���ݰ�꼭�� ÷�ε� ���ڸ��� 1���� ÷�������մϴ�.
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
    
    '÷���� ���ڸ��� �����ڵ�, 121-�ŷ�����, 122-û����, 123-������, 124-���ּ�, 125-�Ա�ǥ,126-������
    SubItemCode = 121
    
    '÷���� ���ڸ��� ������ȣ
    SubMgtKey = "20151223-01"
        
    Set Response = TaxinvoiceService.DetachStatement(txtCorpNum.Text, KeyType, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub




'=========================================================================
' �˺�(www.popbill.com)�� �α��ε� �˺� URL�� ��ȯ�մϴ�.
' - ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ��Ʈ�ʰ����� ��� ��Ʈ�� �ܿ�����Ʈ(GetPartnerBalance API)
'   �� ���� Ȯ���Ͻñ� �ٶ��ϴ�.
'=========================================================================
    
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = TaxinvoiceService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ����ȸ���� ���ڼ��ݰ�꼭 API ���� ���������� Ȯ���մϴ�.
'=========================================================================

Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    
    Set ChargeInfo = TaxinvoiceService.GetChargeInfo(txtCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "unitCost (����ܰ�) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub



'=========================================================================
' ����ȸ���� ȸ�������� Ȯ���մϴ�.
'=========================================================================

Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    
    Set CorpInfo = TaxinvoiceService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "ceoname(��ǥ�ڼ���) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName(��ȣ) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr(�ּ�) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType(����) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass(����) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
    
End Sub

'=========================================================================
' 1���� ���ݰ�꼭 ���׸��� Ȯ���մϴ�.
' - �����׸� ���� �ڼ��� ������ "[���ڼ��ݰ�꼭 API �����Ŵ���]
'   > 4.1 (����)��꼭 ����" �� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================

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
    
    Set tiDetailInfo = TaxinvoiceService.GetDetailInfo(txtCorpNum.Text, KeyType, txtMgtKey.Text)
     
    If tiDetailInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "writeDate (�ۼ�����) : " + tiDetailInfo.writeDate + vbCrLf
    tmp = tmp + "chargeDirection (���ݹ���) : " + tiDetailInfo.chargeDirection + vbCrLf
    tmp = tmp + "issueType (��������) : " + tiDetailInfo.issueType + vbCrLf
    tmp = tmp + "issueTiming (�������) : " + tiDetailInfo.issueTiming + vbCrLf
    tmp = tmp + "taxType (��������) : " + tiDetailInfo.taxType + vbCrLf
    
    tmp = tmp + "invoicerCorpNum (������ ����ڹ�ȣ) : " + tiDetailInfo.invoicerCorpNum + vbCrLf
    tmp = tmp + "invoicerMgtKey (������ ����������ȣ) : " + tiDetailInfo.invoicerMgtKey + vbCrLf
    tmp = tmp + "invoicerTaxRegID (������ ������� �ĺ���ȣ) : " + tiDetailInfo.invoicerTaxRegID + vbCrLf
    tmp = tmp + "invoicerCorpName (������ ��ȣ) : " + tiDetailInfo.invoicerCorpName + vbCrLf
    tmp = tmp + "invoicerCEOName (������ ��ǥ�ڸ�) : " + tiDetailInfo.invoicerCEOName + vbCrLf
    tmp = tmp + "invoicerAddr (������ �ּ�) : " + tiDetailInfo.invoicerAddr + vbCrLf
    tmp = tmp + "invoicerBizClass (������ ����) : " + tiDetailInfo.invoicerBizClass + vbCrLf
    tmp = tmp + "invoicerBizType (������ ����) : " + tiDetailInfo.invoicerBizType + vbCrLf
    tmp = tmp + "invoicerContactName (������ ����ڸ�) : " + tiDetailInfo.invoicerContactName + vbCrLf
    tmp = tmp + "invoicerTEL (������ ����� ����ó) : " + tiDetailInfo.invoicerTEL + vbCrLf
    tmp = tmp + "invoicerHP (������ ����� �޴���) : " + tiDetailInfo.invoicerHP + vbCrLf
    tmp = tmp + "invoicerEmail (������ ����� �̸���) : " + tiDetailInfo.invoicerEmail + vbCrLf
    tmp = tmp + "invoicerSMSSendYN (�ȳ����� ���ۿ���) : " + CStr(tiDetailInfo.invoicerSMSSendYN) + vbCrLf
    
    tmp = tmp + "invoiceeType (���޹޴��� ����) : " + tiDetailInfo.invoiceeType + vbCrLf
    tmp = tmp + "invoiceeCorpNum (���޹޴��� ����ڹ�ȣ) : " + tiDetailInfo.invoiceeCorpNum + vbCrLf
    tmp = tmp + "invoiceeMgtKey (���޹޴��� ����������ȣ) : " + tiDetailInfo.invoiceeMgtKey + vbCrLf
    tmp = tmp + "invoiceeTaxRegID (���޹޴��� ������� �ĺ���ȣ) : " + tiDetailInfo.invoiceeTaxRegID + vbCrLf
    tmp = tmp + "invoiceeCorpName (���޹޴��� ��ȣ) : " + tiDetailInfo.invoiceeCorpName + vbCrLf
    tmp = tmp + "invoiceeCEOName (���޹޴��� ��ǥ�ڼ���) : " + tiDetailInfo.invoiceeCEOName + vbCrLf
    tmp = tmp + "invoiceeAddr (���޹޴��� �ּ�) : " + tiDetailInfo.invoiceeAddr + vbCrLf
    tmp = tmp + "invoiceeBizClass (���޹޴��� ����) : " + tiDetailInfo.invoiceeBizClass + vbCrLf
    tmp = tmp + "invoiceeBizType (���޹޴��� ����) : " + tiDetailInfo.invoiceeBizType + vbCrLf
    tmp = tmp + "invoiceeContactName1 (���޹޴��� ����ڸ�) : " + tiDetailInfo.invoiceeContactName1 + vbCrLf
    tmp = tmp + "invoiceeTEL1 (���޹޴��� ����� ����ó) : " + tiDetailInfo.invoiceeTEL1 + vbCrLf
    tmp = tmp + "invoiceeHP1 (���޹޴��� ����� �޴���) : " + tiDetailInfo.invoiceeHP1 + vbCrLf
    tmp = tmp + "invoiceeEmail1 (���޹޴��� ����� ����) : " + tiDetailInfo.invoiceeEmail1 + vbCrLf
    tmp = tmp + "closeDownState (���޹޴��� ���������) : " + CStr(tiDetailInfo.closeDownState) + vbCrLf
    tmp = tmp + "closeDownStateDate (���޹޴��� ���������) : " + tiDetailInfo.closeDownStateDate + vbCrLf
    
    '''  �󼼳��� ���� '''
    
    MsgBox tmp
    
End Sub

'=========================================================================
' ��뷮 �������� �����ּ� ����� ��ȯ�մϴ�.
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

    For Each email In resultList
        tmp = tmp + email + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ���ݰ�꼭 �μ�(���޹޴���) URL�� ��ȯ�մϴ�.
' - URL ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================

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
    
    url = TaxinvoiceService.GetEPrintURL(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ���ݰ�꼭�� ÷�ε� ������ ����� Ȯ���մϴ�.
' - �����׸� �� ���Ͼ��̵�(AttachedFile) �׸��� ���ϻ���(DeleteFile API)
'   ȣ��� �̿��� �� �ֽ��ϴ�.
'=========================================================================

Private Sub btnGetFiles_Click()
    Dim resultList As Collection
    Dim KeyType As MgtKeyType
    
    Set resultList = TaxinvoiceService.GetFiles(txtCorpNum.Text, KeyType, txtMgtKey.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
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

'=========================================================================
'1���� ���ݰ�꼭 ����/��� ������ Ȯ���մϴ�.
' - ���ݰ�꼭 ��������(GetInfo API) �����׸� ���� �ڼ��� ������
'  "[���ڼ��ݰ�꼭 API �����Ŵ���] > 4.2. (����)��꼭 �������� ����"
'   �� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================

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
    
    Set tiInfo = TaxinvoiceService.GetInfo(txtCorpNum.Text, KeyType, txtMgtKey.Text)
     
    If tiInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "itemKey (���ݰ�꼭 ������Ű) : " + tiInfo.itemKey + vbCrLf
    tmp = tmp + "stateCode (�����ڵ�) : " + CStr(tiInfo.stateCode) + vbCrLf
    tmp = tmp + "taxType (��������) : " + tiInfo.taxType + vbCrLf
    tmp = tmp + "purposeType (����/û��) : " + tiInfo.purposeType + vbCrLf
    tmp = tmp + "modifyCode  (���� �����ڵ�) : " + tiInfo.modifyCode + vbCrLf
    tmp = tmp + "issueType (��������) : " + tiInfo.issueType + vbCrLf
    tmp = tmp + "lateIssueYN (�������� ����) : " + CStr(tiInfo.lateIssueYN) + vbCrLf
    tmp = tmp + "interOPYN (�������� ����) : " + CStr(tiInfo.interOPYN) + vbCrLf
    
    tmp = tmp + "writeDate (�ۼ�����) : " + tiInfo.writeDate + vbCrLf
    
    tmp = tmp + "invoicerCorpName (������ ��ȣ) : " + tiInfo.invoicerCorpName + vbCrLf
    tmp = tmp + "invoicerCorpNum (������ ����ڹ�ȣ) : " + tiInfo.invoicerCorpNum + vbCrLf
    tmp = tmp + "invoicerMgtKey (������ ����������ȣ) : " + tiInfo.invoicerMgtKey + vbCrLf
    tmp = tmp + "invoicerPrintYN (������ �μ⿩��) : " + CStr(tiInfo.invoicerPrintYN) + vbCrLf
    
    tmp = tmp + "invoiceeCorpName (���޹޴��� ��ȣ) : " + tiInfo.invoiceeCorpName + vbCrLf
    tmp = tmp + "invoiceeCorpNum (���޹޴��� ����ڹ�ȣ) : " + tiInfo.invoiceeCorpNum + vbCrLf
    tmp = tmp + "invoiceeMgtKey (���޹޴��� ����������ȣ) : " + tiInfo.invoiceeMgtKey + vbCrLf
    tmp = tmp + "invoiceePrintYN (���޹޴��� �μ⿩��) : " + CStr(tiInfo.invoiceePrintYN) + vbCrLf
    tmp = tmp + "closeDownState (���޹޴��� ���������) : " + CStr(tiInfo.closeDownState) + vbCrLf
    tmp = tmp + "closeDownStateDate (���޹޴��� ���������) : " + tiInfo.closeDownStateDate + vbCrLf
    
    tmp = tmp + "trusteeCorpName (��Ź�� ��ȣ) : " + tiInfo.trusteeCorpName + vbCrLf
    tmp = tmp + "trusteeCorpNum (��Ź�� ����ڹ�ȣ) : " + tiInfo.trusteeCorpNum + vbCrLf
    tmp = tmp + "trusteeMgtKey (��Ź�� ����������ȣ) : " + tiInfo.trusteeMgtKey + vbCrLf
    tmp = tmp + "trusteePrintYN (��Ź�� �μ⿩��) : " + CStr(tiInfo.trusteePrintYN) + vbCrLf
    
    tmp = tmp + "supplyCostTotal (���ް��� �հ�) : " + tiInfo.supplyCostTotal + vbCrLf
    tmp = tmp + "taxTotal (���� �հ�) : " + tiInfo.taxTotal + vbCrLf
    
    tmp = tmp + "issueDT (�����Ͻ�) : " + tiInfo.issueDT + vbCrLf
    tmp = tmp + "preIssueDT (���࿹���Ͻ�) : " + tiInfo.preIssueDT + vbCrLf
    tmp = tmp + "stateDT (���� �����Ͻ�) : " + tiInfo.stateDT + vbCrLf
    tmp = tmp + "openYN (���� ����) : " + CStr(tiInfo.openYN) + vbCrLf
    tmp = tmp + "openDT (���� �Ͻ�) : " + tiInfo.openDT + vbCrLf
    
    tmp = tmp + "ntsresult (����û ���۰��) : " + tiInfo.ntsresult + vbCrLf
    tmp = tmp + "ntsconfirmNum (����û���ι�ȣ) : " + tiInfo.ntsconfirmNum + vbCrLf
    tmp = tmp + "ntssendDT (����û �����Ͻ�) : " + tiInfo.ntssendDT + vbCrLf
    tmp = tmp + "ntsresultDT (����û ��� �����Ͻ�) : " + tiInfo.ntsresultDT + vbCrLf
    tmp = tmp + "ntssendErrCode (���۽��� �����ڵ�) : " + tiInfo.ntssendErrCode + vbCrLf
    
    tmp = tmp + "stateMemo (���¸޸�) : " + tiInfo.stateMemo + vbCrLf
    tmp = tmp + "regDT (�ӽ����� �Ͻ�) : " + tiInfo.regDT + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' �ٷ��� ���ݰ�꼭 ����/��� ������ Ȯ���մϴ�. (�ִ� 1000��)
' - ���ݰ�꼭 ��������(GetInfos API) �����׸� ���� �ڼ��� ������
'  "[���ڼ��ݰ�꼭 API �����Ŵ���] > 4.2. (����)��꼭 �������� ����"
'  �� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnGetInfos_Click()
    Dim resultList As Collection
    Dim KeyType As MgtKeyType
    Dim KeyList As New Collection
    
    KeyType = SELL
    
    '���ݰ�꼭 ����������ȣ �迭, �ִ� 1000��
    KeyList.Add "20161221-03"
    KeyList.Add "20161010-02"
    KeyList.Add "20161010-03"
    KeyList.Add "20161010-04"
    
    Set resultList = TaxinvoiceService.GetInfos(txtCorpNum.Text, KeyType, KeyList)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "ItemKey | StateCode | TaxType | WriteDate | RegDT | InvoicerPrintYN | InvoiceePrintYN | CloseDownState | CloseDownStateDate | InterOPYN" + vbCrLf
    
    Dim info As PBTIInfo
    
    For Each info In resultList
        tmp = tmp + info.itemKey + " | " + CStr(info.stateCode) + " | " + info.taxType + " | " + info.writeDate + " | " + info.regDT + " | "
        tmp = tmp + CStr(info.invoicerPrintYN) + " | " + CStr(info.invoiceePrintYN) + " | " + CStr(info.closeDownState) + " | " + info.closeDownStateDate + " | " + CStr(info.interOPYN) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ���ݰ�꼭 ���� �����̷��� Ȯ���մϴ�.
' - ���� �����̷� Ȯ��(GetLogs API) �����׸� ���� �ڼ��� ������
'   "[���ڼ��ݰ�꼭 API �����Ŵ���] > 3.6.4 ���� �����̷� Ȯ��"
'   �� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnGetLogs_Click()
    Dim resultList As Collection
    Dim KeyType As MgtKeyType
    
    Set resultList = TaxinvoiceService.GetLogs(txtCorpNum.Text, KeyType, txtMgtKey.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
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

'=========================================================================
' ���޹޴��� ���ϸ�ũ URL�� ��ȯ�մϴ�.
' - ���ϸ�ũ URL�� ��ȿ�ð��� �������� �ʽ��ϴ�.
'=========================================================================

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
    
    url = TaxinvoiceService.GetMailURL(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �ټ����� ���ڼ��ݰ�꼭 �μ��˾� URL�� ��ȯ�մϴ�.
' ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
'=========================================================================

Private Sub btnGetMassPrintURL_Click()
    Dim url As String
    Dim KeyType As MgtKeyType
    Dim KeyList As New Collection
    
    '���ڼ��ݰ�꼭 ������ȣ �迭
    KeyType = SELL
    KeyList.Add "20161010-01"
    KeyList.Add "20161010-02"
    KeyList.Add "20161010-03"
    KeyList.Add "20161010-04"
    KeyList.Add "20161010-05"
    
    url = TaxinvoiceService.GetMassPrintURL(txtCorpNum.Text, KeyType, KeyList)
     
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    
End Sub
 
'=========================================================================
' ��Ʈ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ���������� ��� ����ȸ�� �ܿ�����Ʈ(GetBalance API)��
'   �̿��Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = TaxinvoiceService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
    
End Sub

'=========================================================================
' ��Ʈ�� ����Ʈ ���� �˾� URL�� ��ȯ�մϴ�.
' - ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub



'=========================================================================
' 1���� ���ڼ��ݰ�꼭 ���� �˾� URL�� ��ȯ�մϴ�.
' - ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
'=========================================================================

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
    
    url = TaxinvoiceService.GetPopUpURL(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
    
End Sub

'=========================================================================
' 1���� ���ڼ��ݰ�꼭 �μ��˾� URL�� ��ȯ�մϴ�.
' - ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
'=========================================================================

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
    
    url = TaxinvoiceService.GetPrintURL(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �ΰ� �� ÷�ι��� ��� �˾� URL�� ��ȯ�մϴ�.
' - ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
'=========================================================================

Private Sub btnGetSealURL_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetSealURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
    
End Sub



'=========================================================================
' �˺� > ���� ������ �˾� URL�� ��ȯ�մϴ�.
' - ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
'=========================================================================

Private Sub btnGetURL_SBOX_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetURL(txtCorpNum.Text, txtUserID.Text, "SBOX")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �˺� > �ӽ�(����)������ �˾� URL�� ��ȯ�մϴ�.
' - ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
'=========================================================================

Private Sub btnGetURL_TBOX_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetURL(txtCorpNum.Text, txtUserID.Text, "TBOX")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �˺� > ���� �����ۼ� �˾� URL�� ��ȯ�մϴ�.
' - ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
'=========================================================================

Private Sub btnGetURL_WRITE_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetURL(txtCorpNum.Text, txtUserID.Text, "WRITE")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' [�ӽ�����] ������ ���ݰ�꼭�� [����]ó�� �մϴ�.
' - ����(Issue API)�� ȣ���ϴ� �������� ����Ʈ�� �����˴ϴ�.
' - [����Ϸ�] ���ݰ�꼭�� ����ȸ���� ����û ���ۼ����� ����
'   ����/������� ó���˴ϴ�. �⺻����(��������)
' - ����û ���ۼ����� "�˺� �α���" > [���ڼ��ݰ�꼭] > [ȯ�漳��] >
'   [���ڼ��ݰ�꼭 ����] > [����û ���� �� �������� ����] �ǿ���
'   Ȯ���� �� �ֽ��ϴ�.
' - ����û ������å�� ���� ������ "[���ڼ��ݰ�꼭 API �����Ŵ���] >
'   1.4. ����û ���� ��å" �� �����Ͻñ� �ٶ��ϴ�
'=========================================================================

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
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
    
End Sub

'=========================================================================
' ������ ��û���� ���ݰ�꼭�� [����]ó�� �մϴ�.
' - ����(Issue API)�� ȣ���ϴ� �������� ����Ʈ�� �����˴ϴ�.
' - [����Ϸ�] ���ݰ�꼭�� ����ȸ���� ����û ���ۼ����� ����
'    ����/������� ó���˴ϴ�. �⺻����(��������)
' - ����û ���ۼ����� "�˺� �α���" > [���ڼ��ݰ�꼭] > [ȯ�漳��] >
'   [���ڼ��ݰ�꼭 ����] > [����û ���� �� �������� ����] �ǿ���
'   Ȯ���� �� �ֽ��ϴ�.
' - ����û ������å�� ���� ������ "[���ڼ��ݰ�꼭 API �����Ŵ���] >
'   1.4. ����û ���� ��å" �� �����Ͻñ� �ٶ��ϴ�
'=========================================================================

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
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ������ ��û���� ���ݰ�꼭�� [����]ó�� �մϴ�.
' - ����(Issue API)�� ȣ���ϴ� �������� ����Ʈ�� �����˴ϴ�.
' - [����Ϸ�] ���ݰ�꼭�� ����ȸ���� ����û ���ۼ����� ����
'    ����/������� ó���˴ϴ�. �⺻����(��������)
' - ����û ���ۼ����� "�˺� �α���" > [���ڼ��ݰ�꼭] > [ȯ�漳��] >
'   [���ڼ��ݰ�꼭 ����] > [����û ���� �� �������� ����] �ǿ���
'   Ȯ���� �� �ֽ��ϴ�.
' - ����û ������å�� ���� ������ "[���ڼ��ݰ�꼭 API �����Ŵ���] >
'   1.4. ����û ���� ��å" �� �����Ͻñ� �ٶ��ϴ�
'=========================================================================

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
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ��Ʈ���� ����ȸ������ ȸ�������� ��û�մϴ�.
'=========================================================================

Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '��ũ ���̵�
    joinData.LinkID = LinkID
    
    '����ڹ�ȣ, '-'����, 10�ڸ�
    joinData.CorpNum = "1234567890"
    
    '��ǥ�ڼ���, �ִ� 30��
    joinData.ceoname = "��ǥ�ڼ���"
    
    '��ȣ��, �ִ� 70��
    joinData.corpName = "ȸ����ȣ"
    
    '�ּ�, �ִ� 300��
    joinData.addr = "�ּ�"
    
    '����, �ִ� 40��
    joinData.bizType = "����"
    
    '����, �ִ� 40��
    joinData.bizClass = "����"
    
    '���̵�, 6���̻� 20�� �̸�
    joinData.id = "userid"
    
    '��й�ȣ, 6���̻� 20�� �̸�
    joinData.pwd = "pwd_must_be_long_enough"
    
    '����ڸ�, �ִ� 30��
    joinData.ContactName = "����ڼ���"
    
    '����� ����ó, �ִ� 20��
    joinData.ContactTEL = "02-999-9999"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.ContactHP = "010-1234-5678"
    
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.ContactFAX = "02-999-9998"
    
    '����� ����, �ִ� 70��
    joinData.ContactEmail = "test@test.com"
    
    Set Response = TaxinvoiceService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
    
End Sub

'=========================================================================
' ����ȸ���� ����� ����� Ȯ���մϴ�.
'=========================================================================

Private Sub btnListContact_Click()
    Dim resultList As Collection
        
    Set resultList = TaxinvoiceService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "id | email | hp | personName | searchAllAllowYN | tel | fax | mgrYN | regDT | state " + vbCrLf
    
    Dim info As PBContactInfo
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.email + " | " + info.hp + " | " + info.personName + " | " + CStr(info.searchAllAllowYN) _
                + info.tel + " | " + info.fax + " | " + CStr(info.mgrYN) + " | " + info.regDT + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ���ڼ��ݰ�꼭 ���� �������� �׸� ���� ���ۿ��θ� ������� ��ȯ�մϴ�
'=========================================================================
Private Sub btnListemailconfig_Click(index As Integer)
    Dim resultList As Collection
    Dim i As Integer
    
    Set resultList = TaxinvoiceService.ListEmailConfig(txtCorpNum.Text, txtUserID.Text)
    
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
 
    Dim tmp As String
    
    tmp = "������������(EmailType) | ���ۿ���(SendYN) " + vbCrLf
    
    Dim info As PBEmailConfig
    
    For i = 1 To resultList.Count
        If resultList(i).emailType = "TAX_ISSUE" Then
            tmp = tmp + "���޹޴��ڿ��� ���ڼ��ݰ�꼭 ���� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_ISSUE_INVOICER" Then
            tmp = tmp + "�����ڿ��� ���ڼ��ݰ�꼭 ���� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_CHECK" Then
            tmp = tmp + "�����ڿ��� ���ڼ��ݰ�꼭 ����Ȯ�� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_CANCEL_ISSUE" Then
            tmp = tmp + "���޹޴��ڿ��� ���ڼ��ݰ�꼭 ������� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
                    
        If resultList(i).emailType = "TAX_SEND" Then
            tmp = tmp + "���޹޴��ڿ��� [���࿹��] ���ݰ�꼭 �߼� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_ACCEPT" Then
            tmp = tmp + "�����ڿ��� [���࿹��] ���ݰ�꼭 ���� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
            If resultList(i).emailType = "TAX_ACCEPT_ISSUE" Then
            tmp = tmp + "�����ڿ��� [���࿹��] ���ݰ�꼭 �ڵ����� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_DENY" Then
            tmp = tmp + "�����ڿ��� [���࿹��] ���ݰ�꼭 �ź� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If

        If resultList(i).emailType = "TAX_CANCEL_SEND" Then
            tmp = tmp + "���޹޴��ڿ��� [���࿹��] ���ݰ�꼭 ��� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
                    
        If resultList(i).emailType = "TAX_REQUEST" Then
            tmp = tmp + "�����ڿ��� ���ݰ�꼭�� �����û �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_CANCEL_REQUEST" Then
            tmp = tmp + "���޹޴��ڿ��� ���ݰ�꼭 ��� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
            If resultList(i).emailType = "TAX_REFUSE" Then
            tmp = tmp + "���޹޴��ڿ��� ���ݰ�꼭 �ź� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_TRUST_ISSUE" Then
            tmp = tmp + "���޹޴��ڿ��� ���ڼ��ݰ�꼭 ���� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_TRUST_ISSUE_TRUSTEE" Then
            tmp = tmp + "��Ź�ڿ��� ���ڼ��ݰ�꼭 ���� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
                    
        If resultList(i).emailType = "TAX_TRUST_ISSUE_INVOICER" Then
            tmp = tmp + "�����ڿ��� ���ڼ��ݰ�꼭 ���� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_TRUST_CANCEL_ISSUE" Then
            tmp = tmp + "���޹޴��ڿ��� ���ڼ��ݰ�꼭 ������� �˸� : " + resultList(i).emailType + " | "
          tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
    
        If resultList(i).emailType = "TAX_TRUST_SEND" Then
            tmp = tmp + "�����ڿ��� ���ڼ��ݰ�꼭 ������� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_TRUST_CANCEL_ISSUE_INVOICER" Then
            tmp = tmp + "���޹޴��ڿ��� [���࿹��] ���ݰ�꼭 �߼� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_TRUST_ACCEPT" Then
            tmp = tmp + "��Ź�ڿ��� [���࿹��] ���ݰ�꼭 ���� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
                    
        If resultList(i).emailType = "TAX_TRUST_ACCEPT_ISSUE" Then
            tmp = tmp + "��Ź�ڿ��� [���࿹��] ���ݰ�꼭 �ڵ����� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_TRUST_DENY" Then
            tmp = tmp + "��Ź�ڿ��� [���࿹��] ���ݰ�꼭 �ź� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
            If resultList(i).emailType = "TAX_TRUST_CANCEL_SEND" Then
            tmp = tmp + "���޹޴��ڿ��� [���࿹��] ���ݰ�꼭 ��� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_CLOSEDOWN" Then
            tmp = tmp + "�ŷ�ó�� ����� ���� Ȯ�� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_NTSFAIL_INVOICER" Then
            tmp = tmp + "���ڼ��ݰ�꼭 ����û ���۽��� �ȳ�) : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
                    
        If resultList(i).emailType = "TAX_SEND_INFO" Then
            tmp = tmp + "���� �ͼӺ� [���� ���� ���] ���ݰ�꼭 ���� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "ETC_CERT_EXPIRATION" Then
            tmp = tmp + "�˺����� �̿����� ������������ ���� �˸� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
    Next
    
    
    MsgBox tmp

End Sub

'=========================================================================
' ���޹޴��ڿ��� ��û���� ������ ���ݰ�꼭�� [�ź�]ó�� �մϴ�.
' - ���ݰ�꼭�� ����������ȣ�� �����ϱ� ���ؼ��� ���� (Delete API) ��
'   ȣ���Ͽ� [����] ó���ؾ� �մϴ�.
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
    memo = "������ ��û �ź� �޸�"
    
    Set Response = TaxinvoiceService.Refuse(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================================================================================
'[���޹޴���]�� �����ڿ��� 1���� ������ ���ݰ�꼭�� [��� ��û]�մϴ�.
' - ���ݰ�꼭 �׸� ������ "[���ڼ��ݰ�꼭 API �����Ŵ���] > 4.1. (����)��꼭����"�� �����Ͻñ� �ٶ��ϴ�.
' - ������ ���ݰ�꼭 ���μ����� �����ϱ� ���ؼ��� ������/���޹޴��ڰ� ��� �˺��� ȸ���̿��� �մϴ�.
' - ������ ��ÿ�û�� �����ڰ� [����] ó���� ����Ʈ�� �����Ǹ� ������ ���ݰ�꼭 �׸��� ���ݹ���(ChargeDirection)�� ������ ���� ����
'   ������(�����ڰ���) �Ǵ� ������(���޹޴��ڰ���) ó���˴ϴ�.
'=========================================================================================================================================

Private Sub btnRegistRequest_Click()
 Dim Taxinvoice As New PBTaxinvoice
    
   '[�ʼ�] �ۼ�����, ǥ������ (yyyyMMdd) ex)20161010
    Taxinvoice.writeDate = "20181121"
    
    '[�ʼ�] ��������, [������, ������, ����Ź] �� ����
    Taxinvoice.issueType = "������"
    
    '[�ʼ�] {������, ������} �� ����, '������'�� ������ ���μ��������� �̿밡��
    '- ������(������ ����), ������(���޹޴��� ����)
    Taxinvoice.chargeDirection = "������"
    
    '[�ʼ�] ����/û��, [����, û��] �� ����
    Taxinvoice.purposeType = "����"
    
    '[�ʼ�] �������, [��������, ���ν��ڵ�����] �� ����
    ' ���࿹��(Send API) ���μ����� �������� �ʴ°�� '��������' ����
    Taxinvoice.issueTiming = "��������"
    
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
    
    '������ ����������ȣ, 1~24�ڸ� (����, ����, '-', '_') ��������
    '����� ���� �ߺ����� �ʵ��� ����
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
    
    '[������� �ʼ�] ���޹޴��� ����������ȣ(������� �ʼ�)
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
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    '���޹޴��� ����� ����ó
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    
    '���޹޴��� ����� �޴�����ȣ
    Taxinvoice.invoiceeHP1 = "010-111-222"
    
    '������� �����ڿ��� ����ȳ����� ���ۿ���
    Taxinvoice.invoiceeSMSSendYN = False
    
    
    '=========================================================================
    '                            ���ݰ�꼭 ����
    '=========================================================================
    
    '[�ʼ�], ���ް��� �հ�
    Taxinvoice.supplyCostTotal = "200000"
    
    '[�ʼ�] ���� �հ�
    Taxinvoice.taxTotal = "20000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + �����հ�
    Taxinvoice.totalAmount = "220000"
    
    '���� �� '�Ϸù�ȣ' �׸�
    Taxinvoice.serialNum = "123"
    
    '���� �� '��' �׸�, �ִ밪 32767
    Taxinvoice.kwon = "1"
    
    '���� �� 'ȣ' �׸�, �ִ밪 32767
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
    ' - [����] �������ݰ�꼭 �ۼ���� �ȳ� - http://blog.linkhub.co.kr/650
    '========================================================================='
    
    ' ���������ڵ�, ���������� ���� 1~6�� ���ñ���
    Taxinvoice.modifyCode = ""
    
    ' �������ݰ�꼭�� ItemKey, ����Ȯ�� (GetInfo API)�� ������(ItemKey �׸�) Ȯ��
    Taxinvoice.originalTaxinvoiceKey = ""
        
    
    '=========================================================================
    '                            ���׸�(ǰ��) ����
    '=========================================================================
    
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20181120"   '�ŷ�����  yyyyMMdd
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
    newDetail2.purchaseDT = "20181120"   '�ŷ�����  yyyyMMdd
    newDetail2.itemName = "ǰ��2"        'ǰ���
    newDetail2.spec = "�԰�"             '�԰�
    newDetail2.qty = "1"                 '����
    newDetail2.unitCost = "100000"       '�ܰ�
    newDetail2.supplyCost = "100000"     '���ް���
    newDetail2.tax = "10000"             '����
    newDetail2.remark = "���"           '���
    
    Taxinvoice.detailList.Add newDetail2
    
    
    '=========================================================================
    '                            �߰������ ����
    ' - ���ݰ�꼭 ����ȳ� ������ ���Ź��� ���޹޴��� ����ڰ� �ټ��� ���
    ' ����� ������ �߰��Ͽ� ����ȳ������� �ټ����� ������ �� �ֽ��ϴ�.
    '=========================================================================
    
    
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                '�Ϸù�ȣ, 1���� ��������
    newContact.ContactName = "����� ����"  '����ڸ�
    newContact.email = "test2@test.com"     '����� �����ּ�
    Taxinvoice.addContactList.Add newContact
    
    Dim newContact2 As New PBTIContact
    newContact2.serialNum = 2                '�Ϸù�ȣ, 1���� ��������
    newContact2.ContactName = "����� ����"  '����ڸ�
    newContact2.email = "test2@test.com"     '����� �����ּ�
    
    Taxinvoice.addContactList.Add newContact2
        
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

'=========================================================================
' ������ ���ݰ�꼭�� [���] ó���մϴ�.
' - [���]�� ���ݰ�꼭�� ����������ȣ�� �����ϱ� ���ؼ��� ���� (Delete API)
'   �� ȣ���ؾ� �մϴ�.
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
    memo = "������ ��û ��� �޸�"
    
    Set Response = TaxinvoiceService.CancelRequest(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ڼ��ݰ�꼭 ���� �������� �׸� ���� ���ۿ��θ� �����մϴ�.
'=========================================================================
Private Sub btnUpdateemailconfig_Click(index As Integer)
    Dim Response As PBResponse
    Dim emailType As String
    Dim sendYN As Boolean
    
    '���� ���� ����
    emailType = "TAX_ISSUE"

    '���� ���� (True = ����, False = ������)
    sendYN = True
    
    Set Response = TaxinvoiceService.UpdateEmailConfig(txtCorpNum.Text, emailType, sendYN, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���������� ��� URL�� ��ȯ�մϴ�.
' - URL ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetTaxCertURL_Click()

    Dim url As String

    url = TaxinvoiceService.GetTaxCertURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    
    'Internet Explorer Browser ȣ��
    Dim IE As Object
   
    Dim strResult As String
    Dim strSiteName As String
   
    Set IE = CreateObject("InternetExplorer.Application")
    strSiteName = url
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
        .StatusText = "�˺� ���������� ��� �˾� URL"
    End With
   
    Set IE = Nothing
    
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ���� URL�� ��ȯ�մϴ�.
' - URL ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetChargeURL_Click()

    Dim url As String
    
    url = TaxinvoiceService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ���޹޴��ڿ��� ��û���� ������ ���ݰ�꼭�� [�ź�]ó�� �մϴ�.
' - ���ݰ�꼭�� ����������ȣ�� �����ϱ� ���ؼ��� ���� (Delete API) ��
'   ȣ���Ͽ� [����] ó���ؾ� �մϴ�.
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
' ����ȸ���� ����ڸ� �űԷ� ����մϴ�.
'=========================================================================

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�, 6�� �̻� 20�� �̸�
    joinData.id = "testkorea"
    
    '��й�ȣ, 6�� �̻� 20�� �̸�
    joinData.pwd = "test@test.com"
    
    '����ڸ�, �ִ� 30��
    joinData.personName = "����ڸ�"
    
    '����� ����ó
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ
    joinData.hp = "010-1234-1234"
    
    '����� �����ּ�
    joinData.email = "test@test.com"
    
    '����� �ѽ���ȣ
    joinData.fax = "070-1234-1234"
    
    'ȸ����ȸ ���ѿ���, true-ȸ����ȸ / false-������ȸ
    joinData.searchAllAllowYN = True
    
    '������ ���ѿ���
    joinData.mgrYN = False
    
    Set Response = TaxinvoiceService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
    
End Sub

'=========================================================================
' 1���� ���ݰ�꼭�� �ӽ����� �մϴ�.
' - ���ݰ�꼭 �ӽ�����(Register API) ȣ���Ŀ��� ����(Issue API)�� ȣ���ؾ߸�
'   ����û���� ���۵˴ϴ�.
' - �ӽ������ ������ �ѹ��� ȣ��� ó���ϴ� ��ù���(RegistIssue API) ���μ���
'   ������ �����մϴ�.
' - ���ݰ�꼭 �׸� ������ "[���ڼ��ݰ�꼭 API �����Ŵ���] > 4.1. (����)��꼭
'   ����"�� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================
    
Private Sub btnRegister_Click()
    Dim Taxinvoice As New PBTaxinvoice
    Dim writeSpecification As Boolean
    
    '[�ʼ�] �ۼ�����, ǥ������ (yyyyMMdd) ex)20161010
    Taxinvoice.writeDate = "20170223"
    
    '[�ʼ�] ��������, [������, ������, ����Ź] �� ����
    Taxinvoice.issueType = "������"
    
    '[�ʼ�] {������, ������} �� ����, '������'�� ������ ���μ��������� �̿밡��
    '- ������(������ ����), ������(���޹޴��� ����)
    Taxinvoice.chargeDirection = "������"
    
    '[�ʼ�] ����/û��, [����, û��] �� ����
    Taxinvoice.purposeType = "����"
    
    '[�ʼ�] �������, [��������, ���ν��ڵ�����] �� ����
    ' ���࿹��(Send API) ���μ����� �������� �ʴ°�� '��������' ����
    Taxinvoice.issueTiming = "��������"
    
    '[�ʼ�] ��������, [����, ����, �鼼] �� ����
    Taxinvoice.taxType = "����"
    
    
    '=========================================================================
    '                              ������ ����
    '=========================================================================
        
    '[�ʼ�] ������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    Taxinvoice.invoicerCorpNum = "1234567890"
    
    '[�ʼ�] ������ ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    Taxinvoice.invoicerTaxRegID = ""
    
    '[�ʼ�] ������ ��ȣ
    Taxinvoice.invoicerCorpName = "������ ��ȣ"
    
    '[�ʼ�] ������ ����������ȣ, 1~24�ڸ� (����, ����, '-', '_') ��������
    '����� ���� �ߺ����� �ʵ��� ����
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
    
    '������� ���޹޴��ڿ��� ����ȳ����� ���ۿ���
    '- �ȳ����� ���۱�� �̿�� ����Ʈ�� �����˴ϴ�.
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
    
    '[������� �ʼ�] ���޹޴��� ����������ȣ(������� �ʼ�)
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
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    '���޹޴��� ����� ����ó
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    
    '���޹޴��� ����� �޴�����ȣ
    Taxinvoice.invoiceeHP1 = "010-111-222"
    
    '������� �����ڿ��� ����ȳ����� ���ۿ���
    Taxinvoice.invoiceeSMSSendYN = False
            
    
    '=========================================================================
    '                            ���ݰ�꼭 ����
    '=========================================================================
    
    '[�ʼ�], ���ް��� �հ�
    Taxinvoice.supplyCostTotal = "200000"
    
    '[�ʼ�] ���� �հ�
    Taxinvoice.taxTotal = "20000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + �����հ�
    Taxinvoice.totalAmount = "220000"
    
    '���� �� '�Ϸù�ȣ' �׸�
    Taxinvoice.serialNum = "123"
    
    '���� �� '��' �׸�, �ִ밪 32767
    Taxinvoice.kwon = "1"
    
    '���� �� 'ȣ' �׸�, �ִ밪 32767
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
    ' - [����] �������ݰ�꼭 �ۼ���� �ȳ� - http://blog.linkhub.co.kr/650
    '========================================================================='
    
    ' ���������ڵ�, ���������� ���� 1~6�� ���ñ���
    Taxinvoice.modifyCode = ""
    
    ' �������ݰ�꼭�� ItemKey, ����Ȯ�� (GetInfo API)�� ������(ItemKey �׸�) Ȯ��
    Taxinvoice.originalTaxinvoiceKey = ""
        
    
    '=========================================================================
    '                            ���׸�(ǰ��) ����
    '=========================================================================
    
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20161010"   '�ŷ�����  yyyyMMdd
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
    newDetail2.purchaseDT = "20161010"   '�ŷ�����  yyyyMMdd
    newDetail2.itemName = "ǰ��2"        'ǰ���
    newDetail2.spec = "�԰�"             '�԰�
    newDetail2.qty = "1"                 '����
    newDetail2.unitCost = "100000"       '�ܰ�
    newDetail2.supplyCost = "100000"     '���ް���
    newDetail2.tax = "10000"             '����
    newDetail2.remark = "���"           '���
    
    Taxinvoice.detailList.Add newDetail2
    
    
    '=========================================================================
    '                              �߰������ ����
    ' - ���ݰ�꼭 ����ȳ� ������ ���Ź��� ���޹޴��� ����ڰ� �ټ��� ���
    ' ����� ������ �߰��Ͽ� ����ȳ������� �ټ����� ������ �� �ֽ��ϴ�.
    '=========================================================================
    
    
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                '�Ϸù�ȣ, 1���� ��������
    newContact.ContactName = "����� ����"  '����ڸ�
    newContact.email = "test2@test.com"     '����� �����ּ�
    Taxinvoice.addContactList.Add newContact
    
    Dim newContact2 As New PBTIContact
    newContact2.serialNum = 2                '�Ϸù�ȣ, 1���� ��������
    newContact2.ContactName = "����� ����"  '����ڸ�
    newContact2.email = "test2@test.com"     '����� �����ּ�
    
    Taxinvoice.addContactList.Add newContact2
    
    '�ŷ����� �����ۼ� ����
    writeSpecification = True
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.Register(txtCorpNum.Text, Taxinvoice, writeSpecification)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
    

End Sub

'=========================================================================
' 1���� ������ ���ݰ�꼭�� [�ӽ�����] �մϴ�.
' - ���ݰ�꼭 �׸� ������ "[���ڼ��ݰ�꼭 API �����Ŵ���] > 4.1. (����)��꼭
'   ����"�� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnRegister_rev_Click()
    Dim Taxinvoice As New PBTaxinvoice
    
    '[�ʼ�] �ۼ�����, ǥ������ (yyyyMMdd) ex)20161010
    Taxinvoice.writeDate = "20170223"
    
    '[�ʼ�] ��������, [������, ������, ����Ź] �� ����
    Taxinvoice.issueType = "������"
    
    '[�ʼ�] {������, ������} �� ����, '������'�� ������ ���μ��������� �̿밡��
    '- ������(������ ����), ������(���޹޴��� ����)
    Taxinvoice.chargeDirection = "������"
    
    '[�ʼ�] ����/û��, [����, û��] �� ����
    Taxinvoice.purposeType = "����"
    
    '[�ʼ�] �������, [��������, ���ν��ڵ�����] �� ����
    ' ���࿹��(Send API) ���μ����� �������� �ʴ°�� '��������' ����
    Taxinvoice.issueTiming = "��������"
    
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
    
    '[�ʼ�] ������ ����������ȣ, 1~24�ڸ� (����, ����, '-', '_') ��������
    '����� ���� �ߺ����� �ʵ��� ����
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
    Taxinvoice.invoiceeCorpNum = "1234567890"
    
    '[�ʼ�] ���޹޴��� ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    Taxinvoice.invoiceeTaxRegID = ""
    
    '[�ʼ�] �����ڹ޴��� ��ȣ
    Taxinvoice.invoiceeCorpName = "���޹޴��� ��ȣ"
    
    '[������� �ʼ�] ���޹޴��� ����������ȣ(������� �ʼ�)
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
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    '���޹޴��� ����� ����ó
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    
    '���޹޴��� ����� �޴�����ȣ
    Taxinvoice.invoiceeHP1 = "010-111-222"
    
    '������� �����ڿ��� ����ȳ����� ���ۿ���
    Taxinvoice.invoiceeSMSSendYN = False
            
    
    '=========================================================================
    '                            ���ݰ�꼭 ����
    '=========================================================================
    
    '[�ʼ�], ���ް��� �հ�
    Taxinvoice.supplyCostTotal = "200000"
    
    '[�ʼ�] ���� �հ�
    Taxinvoice.taxTotal = "20000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + �����հ�
    Taxinvoice.totalAmount = "220000"
    
    '���� �� '�Ϸù�ȣ' �׸�
    Taxinvoice.serialNum = "123"
    
    '���� �� '��' �׸�, �ִ밪 32767
    Taxinvoice.kwon = "1"
    
    '���� �� 'ȣ' �׸�, �ִ밪 32767
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
    ' - [����] �������ݰ�꼭 �ۼ���� �ȳ� - http://blog.linkhub.co.kr/650
    '========================================================================='
    
    ' ���������ڵ�, ���������� ���� 1~6�� ���ñ���
    Taxinvoice.modifyCode = ""
    
    ' �������ݰ�꼭�� ItemKey, ����Ȯ�� (GetInfo API)�� ������(ItemKey �׸�) Ȯ��
    Taxinvoice.originalTaxinvoiceKey = ""
        
    
    '=========================================================================
    '                            ���׸�(ǰ��) ����
    '=========================================================================
    
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20161010"   '�ŷ�����  yyyyMMdd
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
    newDetail2.purchaseDT = "20161010"   '�ŷ�����  yyyyMMdd
    newDetail2.itemName = "ǰ��2"        'ǰ���
    newDetail2.spec = "�԰�"             '�԰�
    newDetail2.qty = "1"                 '����
    newDetail2.unitCost = "100000"       '�ܰ�
    newDetail2.supplyCost = "100000"     '���ް���
    newDetail2.tax = "10000"             '����
    newDetail2.remark = "���"           '���
    
    Taxinvoice.detailList.Add newDetail2
    
    
    '=========================================================================
    '                              �߰������ ����
    ' - ���ݰ�꼭 ����ȳ� ������ ���Ź��� ���޹޴��� ����ڰ� �ټ��� ���
    ' ����� ������ �߰��Ͽ� ����ȳ������� �ټ����� ������ �� �ֽ��ϴ�.
    '=========================================================================
    
    
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                '�Ϸù�ȣ, 1���� ��������
    newContact.ContactName = "����� ����"  '����ڸ�
    newContact.email = "test2@test.com"     '����� �����ּ�
    Taxinvoice.addContactList.Add newContact
    
    Dim newContact2 As New PBTIContact
    newContact2.serialNum = 2                '�Ϸù�ȣ, 1���� ��������
    newContact2.ContactName = "����� ����"  '����ڸ�
    newContact2.email = "test2@test.com"     '����� �����ּ�
    
    Taxinvoice.addContactList.Add newContact2
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.Register(txtCorpNum.Text, Taxinvoice, False)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' 1���� ���ݰ�꼭�� ��ù��� ó���մϴ�.
' - ���ݰ�꼭 �׸� ������ "[���ڼ��ݰ�꼭 API �����Ŵ���] > 4.1. (����)��꼭
'   ����"�� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnRegistIssue_Click()
    Dim Taxinvoice As New PBTaxinvoice
        
   '[�ʼ�] �ۼ�����, ǥ������ (yyyyMMdd) ex)20161010
    Taxinvoice.writeDate = "20180911"
    
    '[�ʼ�] ��������, [������, ������, ����Ź] �� ����
    Taxinvoice.issueType = "������"
    
    '[�ʼ�] {������, ������} �� ����, '������'�� ������ ���μ��������� �̿밡��
    '- ������(������ ����), ������(���޹޴��� ����)
    Taxinvoice.chargeDirection = "������"
    
    '[�ʼ�] ����/û��, [����, û��] �� ����
    Taxinvoice.purposeType = "����"
    
    '[�ʼ�] �������, [��������, ���ν��ڵ�����] �� ����
    ' ���࿹��(Send API) ���μ����� �������� �ʴ°�� '��������' ����
    Taxinvoice.issueTiming = "��������"
    
    '[�ʼ�] ��������, [����, ����, �鼼] �� ����
    Taxinvoice.taxType = "����"
    
    
    '=========================================================================
    '                              ������ ����
    '=========================================================================
        
    '[�ʼ�] ������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    Taxinvoice.invoicerCorpNum = "1234567890"
    
    '[�ʼ�] ������ ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    Taxinvoice.invoicerTaxRegID = ""
    
    '[�ʼ�] ������ ��ȣ
    Taxinvoice.invoicerCorpName = "������ ��ȣ"
    
    '[�ʼ�] ������ ����������ȣ, 1~24�ڸ� (����, ����, '-', '_') ��������
    '����� ���� �ߺ����� �ʵ��� ����
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
    
    '������� ���޹޴��ڿ��� ����ȳ����� ���ۿ���
    '- �ȳ����� ���۱�� �̿�� ����Ʈ�� �����˴ϴ�.
    Taxinvoice.invoicerSMSSendYN = False
    
    
    '=========================================================================
    '                            ���޹޴��� ����
    '=========================================================================
        
    '[�ʼ�] ���޹޴��� ����, [�����, ����, �ܱ���] �� ����
    Taxinvoice.invoiceeType = "�����"
    
    '[�ʼ�] ���޹޴��� ����ڹ�ȣ, '-' ���� 10�ڸ�
    Taxinvoice.invoiceeCorpNum = "2258300280"
    
    '[�ʼ�] ���޹޴��� ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    Taxinvoice.invoiceeTaxRegID = ""
    
    '[�ʼ�] �����ڹ޴��� ��ȣ
    Taxinvoice.invoiceeCorpName = "���޹޴��� ��ȣ"
    
    '[������� �ʼ�] ���޹޴��� ����������ȣ(������� �ʼ�)
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
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    '���޹޴��� ����� ����ó
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    
    '���޹޴��� ����� �޴�����ȣ
    Taxinvoice.invoiceeHP1 = "010-111-222"
    
    '������� �����ڿ��� ����ȳ����� ���ۿ���
    Taxinvoice.invoiceeSMSSendYN = False
            
    
    '=========================================================================
    '                            ���ݰ�꼭 ����
    '=========================================================================
    
    '[�ʼ�], ���ް��� �հ�
    Taxinvoice.supplyCostTotal = "200000"
    
    '[�ʼ�] ���� �հ�
    Taxinvoice.taxTotal = "20000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + �����հ�
    Taxinvoice.totalAmount = "220000"
    
    '���� �� '�Ϸù�ȣ' �׸�
    Taxinvoice.serialNum = "123"
    
    '���� �� '��' �׸�, �ִ밪 32767
    Taxinvoice.kwon = "1"
    
    '���� �� 'ȣ' �׸�, �ִ밪 32767
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
    ' - [����] �������ݰ�꼭 �ۼ���� �ȳ� - http://blog.linkhub.co.kr/650
    '========================================================================='
    
    ' ���������ڵ�, ���������� ���� 1~6�� ���ñ���
    Taxinvoice.modifyCode = ""
    
    ' �������ݰ�꼭�� ItemKey, ����Ȯ�� (GetInfo API)�� ������(ItemKey �׸�) Ȯ��
    Taxinvoice.originalTaxinvoiceKey = ""
        
    
    '=========================================================================
    '                            ���׸�(ǰ��) ����
    '=========================================================================
    
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20161010"   '�ŷ�����  yyyyMMdd
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
    newDetail2.purchaseDT = "20161010"   '�ŷ�����  yyyyMMdd
    newDetail2.itemName = "ǰ��2"        'ǰ���
    newDetail2.spec = "�԰�"             '�԰�
    newDetail2.qty = "1"                 '����
    newDetail2.unitCost = "100000"       '�ܰ�
    newDetail2.supplyCost = "100000"     '���ް���
    newDetail2.tax = "10000"             '����
    newDetail2.remark = "���"           '���
    
    Taxinvoice.detailList.Add newDetail2
    
    
    '=========================================================================
    '                              �߰������ ����
    ' - ���ݰ�꼭 ����ȳ� ������ ���Ź��� ���޹޴��� ����ڰ� �ټ��� ���
    ' ����� ������ �߰��Ͽ� ����ȳ������� �ټ����� ������ �� �ֽ��ϴ�.
    '=========================================================================
    
    
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                '�Ϸù�ȣ, 1���� ��������
    newContact.ContactName = "����� ����"  '����ڸ�
    newContact.email = "test2@test.com"     '����� �����ּ�
    Taxinvoice.addContactList.Add newContact
    
    Dim newContact2 As New PBTIContact
    newContact2.serialNum = 2                '�Ϸù�ȣ, 1���� ��������
    newContact2.ContactName = "����� ����"  '����ڸ�
    newContact2.email = "test2@test.com"     '����� �����ּ�
    
    Taxinvoice.addContactList.Add newContact2
        
        
    '�ŷ����� �����ۼ� ����
    Taxinvoice.writeSpecification = False
    
    '�������� ��������(forceIssue)
    '���ึ������ ���� ���ݰ�꼭�� �����ϴ� ���, ���꼼�� �ΰ��� �� �ֽ��ϴ�.
    '���꼼�� �ΰ��Ǵ��� ������ �ؾ��ϴ� ��쿡�� forceIssue�� ����
    'true�� �����Ͽ� ����(Issue API)�� ȣ���Ͻø� �˴ϴ�.
    Taxinvoice.forceIssue = False
    
    '�޸�
    Taxinvoice.memo = ""
    
    '����ȳ� ��������, ����ó���� �⺻�������� ����
    Taxinvoice.emailSubject = ""
    
    '�ŷ����� �����ۼ��� �ŷ����� ������ȣ, �̱���� ���ݰ�꼭 ������ȣ�� �ڵ��ۼ�
    Taxinvoice.dealInvoiceMgtKey = ""
    
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.RegistIssue(txtCorpNum.Text, Taxinvoice)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
    
End Sub

'=========================================================================
' ���޹޴��ڰ� �����ڿ��� 1���� ������ ���ݰ�꼭�� ��û�մϴ�.
' - ������ ���ݰ�꼭 ���μ����� �����ϱ� ���ؼ��� ������/���޹޴��ڰ� ���
'   �˺��� ȸ���̿��� �մϴ�.
' - ������ ��û�� �����ڰ� [����] ó���� ����Ʈ�� �����Ǹ� ������
'   ���ݰ�꼭 �׸��� ���ݹ���(ChargeDirection) �� ������ ���� ����
'   ������(�����ڰ���) �Ǵ� ������(���޹޴��ڰ���) ó���˴ϴ�.
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

'=========================================================================
' ������ ���ݰ�꼭�� [���] ó���մϴ�.
' - [���]�� ���ݰ�꼭�� ����������ȣ�� �����ϱ� ���ؼ��� ���� (Delete API)
'   �� ȣ���ؾ� �մϴ�.
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
    memo = "������ ��û ��� �޸�"
    
    Set Response = TaxinvoiceService.CancelRequest(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �˻������� ����Ͽ� ���ݰ�꼭 ����� ��ȸ�մϴ�.
' - �����׸� ���� �ڼ��� ������ "[���ڼ��ݰ�꼭 API �����Ŵ���] >
'   4.2. (����)��꼭 �������� ����" �� �����Ͻñ� �ٶ��ϴ�.
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
    Dim LateOnly As String
    Dim Page As Integer
    Dim PerPage As Integer
    Dim Order As String
    Dim TaxRegIDType As String
    Dim TaxRegID As String
    Dim TaxRegIDYN As String
    Dim QString As String
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
    SDate = "20171101"
    
    '[�ʼ�] ��������, yyyyMMdd
    EDate = "20171231"
    
    '���ۻ��°� �迭, �̱���� ��ü������ȸ, �������°� 3�ڸ����� �ۼ�
    '2,3��° ���ϵ�ī�� ����
    state.Add "3**"
    state.Add "6**"
    
    '�������� �迭, N-�Ϲ� M-���� �� ����, �̱���� ��ü��ȸ
    TType.Add "N"
    TType.Add "M"
    
    '�������� �迭, T-����, N-�鼼 Z-���� �� ����, �̱���� ��ü��ȸ
    taxType.Add "T"
    taxType.Add "N"
    taxType.Add "Z"
    
    '�������� �迭, N-������, R-������, T-����Ź
    issueType.Add "N"
    issueType.Add "R"
    issueType.Add "T"
    
    '�������� ����, 0-�������и� ��ȸ 1-��������и���ȸ, ����ó���� ��ü��ȸ
    LateOnly = ""
    
    '������ ��ȣ
    Page = 1
    
    '������ ��ϰ���, �ִ� 1000��
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
    
    '�������� ����, ����-��ü��ȸ, 0-�Ϲݹ��� ��ȸ, 1-�������� ��ȸ
    interOPYN = ""
    
    Set tiSearchList = TaxinvoiceService.Search(txtCorpNum.Text, KeyType, DType, SDate, EDate, state, TType, _
                        taxType, LateOnly, Page, PerPage, Order, TaxRegIDType, TaxRegID, TaxRegIDYN, QString, _
                        txtUserID.Text, interOPYN, issueType)
     
    If tiSearchList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    tmp = "code (�����ڵ�) : " + CStr(tiSearchList.code) + vbCrLf
    tmp = tmp + "total (�� �˻���� �Ǽ�) : " + CStr(tiSearchList.total) + vbCrLf
    tmp = tmp + "perPage (�������� �˻�����) : " + CStr(tiSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum (������ ��ȣ) : " + CStr(tiSearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (������ ����) : " + CStr(tiSearchList.pageCount) + vbCrLf
    tmp = tmp + "message (����޽���) : " + tiSearchList.message + vbCrLf + vbCrLf
                
    Dim tiInfo As PBTIInfo
    
    For Each tiInfo In tiSearchList.list
        tmp = tmp + "itemKey (���ݰ�꼭 ������Ű) : " + tiInfo.itemKey + vbCrLf
        tmp = tmp + "taxType (��������) : " + tiInfo.taxType + vbCrLf
        tmp = tmp + "writeDate (�ۼ�����) : " + tiInfo.writeDate + vbCrLf
        tmp = tmp + "regDT (�ӽ����� ����) : " + tiInfo.regDT + vbCrLf
        tmp = tmp + "issueType (��������) : " + tiInfo.issueType + vbCrLf
        tmp = tmp + "supplyCostTotal (���ް��� �հ�) : " + tiInfo.supplyCostTotal + vbCrLf
        tmp = tmp + "taxTotal (���� �հ�) : " + tiInfo.taxTotal + vbCrLf
        tmp = tmp + "purposeType (����/û��) : " + tiInfo.purposeType + vbCrLf
        tmp = tmp + "issueDT (�����Ͻ�) : " + tiInfo.issueDT + vbCrLf
        tmp = tmp + "lateIssueYN (�������� ����) : " + CStr(tiInfo.lateIssueYN) + vbCrLf
        tmp = tmp + "preIssueDT (���࿹���Ͻ�) : " + tiInfo.preIssueDT + vbCrLf
        tmp = tmp + "openYN (���� ����) : " + CStr(tiInfo.openYN) + vbCrLf
        tmp = tmp + "openDT (���� �Ͻ�) : " + tiInfo.openDT + vbCrLf
        tmp = tmp + "stateMemo (���¸޸�) : " + tiInfo.stateMemo + vbCrLf
        tmp = tmp + "stateCode (�����ڵ�) : " + CStr(tiInfo.stateCode) + vbCrLf
        
        tmp = tmp + "modifyCode  (���� �����ڵ�) : " + tiInfo.modifyCode + vbCrLf
        
        tmp = tmp + "interOPYN (�������� ����) : " + CStr(tiInfo.interOPYN) + vbCrLf
        
        tmp = tmp + "invoicerCorpName (������ ��ȣ) : " + tiInfo.invoicerCorpName + vbCrLf
        tmp = tmp + "invoicerCorpNum (������ ����ڹ�ȣ) : " + tiInfo.invoicerCorpNum + vbCrLf
        tmp = tmp + "invoicerMgtKey (������ ����������ȣ) : " + tiInfo.invoicerMgtKey + vbCrLf
        tmp = tmp + "invoicerPrintYN (������ �μ⿩��) : " + CStr(tiInfo.invoicerPrintYN) + vbCrLf
        
        tmp = tmp + "invoiceeCorpName (���޹޴��� ��ȣ) : " + tiInfo.invoiceeCorpName + vbCrLf
        tmp = tmp + "invoiceeCorpNum (���޹޴��� ����ڹ�ȣ) : " + tiInfo.invoiceeCorpNum + vbCrLf
        tmp = tmp + "invoiceeMgtKey (���޹޴��� ����������ȣ) : " + tiInfo.invoiceeMgtKey + vbCrLf
        tmp = tmp + "invoiceePrintYN (���޹޴��� �μ⿩��) : " + CStr(tiInfo.invoiceePrintYN) + vbCrLf
        tmp = tmp + "closeDownState (���޹޴��� ���������) : " + CStr(tiInfo.closeDownState) + vbCrLf
        tmp = tmp + "closeDownStateDate (���޹޴��� ���������) : " + tiInfo.closeDownStateDate + vbCrLf
        
        tmp = tmp + "trusteeCorpName (��Ź�� ��ȣ) : " + tiInfo.trusteeCorpName + vbCrLf
        tmp = tmp + "trusteeCorpNum (��Ź�� ����ڹ�ȣ) : " + tiInfo.trusteeCorpNum + vbCrLf
        tmp = tmp + "trusteeMgtKey (��Ź�� ����������ȣ) : " + tiInfo.trusteeMgtKey + vbCrLf
        tmp = tmp + "trusteePrintYN (��Ź�� �μ⿩��) : " + CStr(tiInfo.trusteePrintYN) + vbCrLf
        
        
        tmp = tmp + "stateDT (���� �����Ͻ�) : " + tiInfo.stateDT + vbCrLf
        
        
        tmp = tmp + "ntsresult (����û ���۰��) : " + tiInfo.ntsresult + vbCrLf
        tmp = tmp + "ntsconfirmNum (����û���ι�ȣ) : " + tiInfo.ntsconfirmNum + vbCrLf
        tmp = tmp + "ntssendDT (����û �����Ͻ�) : " + tiInfo.ntssendDT + vbCrLf
        tmp = tmp + "ntsresultDT (����û ��� �����Ͻ�) : " + tiInfo.ntsresultDT + vbCrLf
        tmp = tmp + "ntssendErrCode (���۽��� �����ڵ�) : " + tiInfo.ntssendErrCode + vbCrLf
        
        
        
    Next
    
    MsgBox tmp
       
End Sub

'=========================================================================
' 1���� [�ӽ�����] ������ ���ݰ�꼭�� [���࿹��] ó���մϴ�.
' - ���࿹���̶� �����ڿ� ���޹޴��� ���̿� ���ݰ�꼭 Ȯ�� �� �����ϴ�
'   ����Դϴ�.
' - "[���ڼ��ݰ�꼭 API �����Ŵ���] > 1.3.1. ������ ���μ��� �帧��
'   > ��. �ӽ����� ���࿹��" �� ���μ����� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================

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
    
    '���࿹�� �ȳ����� ����, ����ó���� �⺻�������� ����
    emailSubject = ""
    
    '�޸�
    memo = "���࿹�� �޸�"
    
    Set Response = TaxinvoiceService.Send(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, emailSubject)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���� �ȳ������� �������մϴ�.
' - ���ϳ����� ���ڼ��ݰ�꼭 [����] ��ư�� �������� �ʴ� ���,
'   Ű���� ���� Shift Ű�� ������ ��ư�� Ŭ���غ��ñ� �ٶ��ϴ�.
'=========================================================================

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
    
    '������ �����ּ�
    receiveEmail = "test@test.com"
    
    Set Response = TaxinvoiceService.SendEmail(txtCorpNum.Text, KeyType, txtMgtKey.Text, receiveEmail)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ڼ��ݰ�꼭�� �ѽ������մϴ�.
' - �ѽ� ���� ��û�� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
' - ���۳��� Ȯ���� "�˺� �α���" > [���� �ѽ�] > [�ѽ�] > [���۳���]
'   �޴����� ���۰���� Ȯ���� �� �ֽ��ϴ�.
'=========================================================================

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
        
    '�߽Ź�ȣ
    senderNum = "07043042991"
    
    '���Ź�ȣ
    receiveNum = "111-222-4444"
        
    Set Response = TaxinvoiceService.SendFax(txtCorpNum.Text, KeyType, txtMgtKey.Text, senderNum, receiveNum)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �˸����ڸ� �����մϴ�. (�ܹ�/SMS- �ѱ� �ִ� 45��)
' - �˸����� ���۽� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
' - ���۳��� Ȯ���� "�˺� �α���" > [���� �ѽ�] > [���۳���] �ǿ���
'   ���۰���� Ȯ���� �� �ֽ��ϴ�.
'=========================================================================
        
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
    
    ' �߽Ź�ȣ, [����] �߽Ź�ȣ ��Ģ���� - http://blog.linkhub.co.kr/3064
    senderNum = "07043042991"
    
    ' ���Ź�ȣ
    receiveNum = "010-111-222"
    
    ' �޽��� ����, �ִ� 90Byte(�ѱ� 45��), ���̸� �ʰ��� ������ �����Ǿ� ���۵˴ϴ�.
    Contents = "��ũ��꿡�� ���ݰ�꼭�� �����Ͽ����ϴ�. ����Ȯ�� �ٶ��ϴ�."
    
    Set Response = TaxinvoiceService.SendSMS(txtCorpNum.Text, KeyType, txtMgtKey.Text, _
                                senderNum, receiveNum, Contents)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub


'=========================================================================
' [����Ϸ�] ������ ���ݰ�꼭�� ����û���� ��������մϴ�.
' - ����û ��������� ȣ������ ���� ���ݰ�꼭�� ������ ���� ���� ���� 3�ÿ�
'   �˺� �ý��ۿ��� �ϰ������� ����û���� �����մϴ�.
' - �������۽� �������� ������������ ��� ���� �����Ͽ� ���۵˴ϴ�.
' - ����û ���ۿ� ���� ������ "[���ڼ��ݰ�꼭 API �����Ŵ���] > 1.4 ����û
'   ���� ��å" �� �����Ͻñ� �ٶ��ϴ�.
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
    
    Set Response = TaxinvoiceService.SendToNTS(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ڼ��ݰ�꼭 ����ܰ��� Ȯ���մϴ�.
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
' [�ӽ�����] ������ ���ݰ�꼭�� �׸��� �����մϴ�.
' - ���ݰ�꼭 �׸� ������ "[���ڼ��ݰ�꼭 API �����Ŵ���] > 4.1. (����)��꼭
'   ����"�� �����Ͻñ� �ٶ��ϴ�.
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
    
    '[�ʼ�] �ۼ�����, ǥ������ (yyyyMMdd) ex)20161010
    Taxinvoice.writeDate = "20170223"
    
    '[�ʼ�] ��������, [������, ������, ����Ź] �� ����
    Taxinvoice.issueType = "������"
    
    '[�ʼ�] {������, ������} �� ����, '������'�� ������ ���μ��������� �̿밡��
    '- ������(������ ����), ������(���޹޴��� ����)
    Taxinvoice.chargeDirection = "������"
    
    '[�ʼ�] ����/û��, [����, û��] �� ����
    Taxinvoice.purposeType = "����"
    
    '[�ʼ�] �������, [��������, ���ν��ڵ�����] �� ����
    ' ���࿹��(Send API) ���μ����� �������� �ʴ°�� '��������' ����
    Taxinvoice.issueTiming = "��������"
    
    '[�ʼ�] ��������, [����, ����, �鼼] �� ����
    Taxinvoice.taxType = "����"
    
    
    '=========================================================================
    '                              ������ ����
    '=========================================================================
        
    '[�ʼ�] ������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    Taxinvoice.invoicerCorpNum = "1234567890"
    
    '[�ʼ�] ������ ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    Taxinvoice.invoicerTaxRegID = ""
    
    '[�ʼ�] ������ ��ȣ
    Taxinvoice.invoicerCorpName = "������ ��ȣ_����"
    
    '[�ʼ�] ������ ����������ȣ, 1~24�ڸ� (����, ����, '-', '_') ��������
    '����� ���� �ߺ����� �ʵ��� ����
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
    
    '������� ���޹޴��ڿ��� ����ȳ����� ���ۿ���
    '- �ȳ����� ���۱�� �̿�� ����Ʈ�� �����˴ϴ�.
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
    
    '[������� �ʼ�] ���޹޴��� ����������ȣ(������� �ʼ�)
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
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    '���޹޴��� ����� ����ó
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    
    '���޹޴��� ����� �޴�����ȣ
    Taxinvoice.invoiceeHP1 = "010-111-222"
    
    '������� �����ڿ��� ����ȳ����� ���ۿ���
    Taxinvoice.invoiceeSMSSendYN = False
            
    
    '=========================================================================
    '                            ���ݰ�꼭 ����
    '=========================================================================
    
    '[�ʼ�], ���ް��� �հ�
    Taxinvoice.supplyCostTotal = "200000"
    
    '[�ʼ�] ���� �հ�
    Taxinvoice.taxTotal = "20000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + �����հ�
    Taxinvoice.totalAmount = "220000"
    
    '���� �� '�Ϸù�ȣ' �׸�
    Taxinvoice.serialNum = "123"
    
    '���� �� '��' �׸�, �ִ밪 32767
    Taxinvoice.kwon = "1"
    
    '���� �� 'ȣ' �׸�, �ִ밪 32767
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
    ' - [����] �������ݰ�꼭 �ۼ���� �ȳ� - http://blog.linkhub.co.kr/650
    '========================================================================='
    
    ' ���������ڵ�, ���������� ���� 1~6�� ���ñ���
    Taxinvoice.modifyCode = ""
    
    ' �������ݰ�꼭�� ItemKey, ����Ȯ�� (GetInfo API)�� ������(ItemKey �׸�) Ȯ��
    Taxinvoice.originalTaxinvoiceKey = ""
        
    
    '=========================================================================
    '                            ���׸�(ǰ��) ����
    '=========================================================================
    
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20161010"   '�ŷ�����  yyyyMMdd
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
    newDetail2.purchaseDT = "20161010"   '�ŷ�����  yyyyMMdd
    newDetail2.itemName = "ǰ��2"        'ǰ���
    newDetail2.spec = "�԰�"             '�԰�
    newDetail2.qty = "1"                 '����
    newDetail2.unitCost = "100000"       '�ܰ�
    newDetail2.supplyCost = "100000"     '���ް���
    newDetail2.tax = "10000"             '����
    newDetail2.remark = "���"           '���
    
    Taxinvoice.detailList.Add newDetail2
    
    
    '=========================================================================
    '                              �߰������ ����
    ' - ���ݰ�꼭 ����ȳ� ������ ���Ź��� ���޹޴��� ����ڰ� �ټ��� ���
    ' ����� ������ �߰��Ͽ� ����ȳ������� �ټ����� ������ �� �ֽ��ϴ�.
    '=========================================================================
    
    
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                '�Ϸù�ȣ, 1���� ��������
    newContact.ContactName = "����� ����"  '����ڸ�
    newContact.email = "test2@test.com"     '����� �����ּ�
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

'=========================================================================
' [�ӽ�����] ������ ���ݰ�꼭�� �׸��� �����մϴ�.
' - ���ݰ�꼭 �׸� ������ "[���ڼ��ݰ�꼭 API �����Ŵ���] > 4.1. (����)��꼭
'   ����"�� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnUpdate_rev_Click()
    Dim KeyType As MgtKeyType
    
    KeyType = BUY
    
    Dim Taxinvoice As New PBTaxinvoice
    
    '[�ʼ�] �ۼ�����, ǥ������ (yyyyMMdd) ex)20161010
    Taxinvoice.writeDate = "20170223"
    
    '[�ʼ�] ��������, [������, ������, ����Ź] �� ����
    Taxinvoice.issueType = "������"
    
    '[�ʼ�] {������, ������} �� ����, '������'�� ������ ���μ��������� �̿밡��
    '- ������(������ ����), ������(���޹޴��� ����)
    Taxinvoice.chargeDirection = "������"
    
    '[�ʼ�] ����/û��, [����, û��] �� ����
    Taxinvoice.purposeType = "����"
    
    '[�ʼ�] �������, [��������, ���ν��ڵ�����] �� ����
    ' ���࿹��(Send API) ���μ����� �������� �ʴ°�� '��������' ����
    Taxinvoice.issueTiming = "��������"
    
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
    
    '[�ʼ�] ������ ����������ȣ, 1~24�ڸ� (����, ����, '-', '_') ��������
    '����� ���� �ߺ����� �ʵ��� ����
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
    
    '������� ���޹޴��ڿ��� ����ȳ����� ���ۿ���
    '- �ȳ����� ���۱�� �̿�� ����Ʈ�� �����˴ϴ�.
    Taxinvoice.invoicerSMSSendYN = False
    
    
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
    
    '[������� �ʼ�] ���޹޴��� ����������ȣ(������� �ʼ�)
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
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    '���޹޴��� ����� ����ó
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    
    '���޹޴��� ����� �޴�����ȣ
    Taxinvoice.invoiceeHP1 = "010-111-222"
    
    '������� �����ڿ��� ����ȳ����� ���ۿ���
    Taxinvoice.invoiceeSMSSendYN = False
            
    
    '=========================================================================
    '                            ���ݰ�꼭 ����
    '=========================================================================
    
    '[�ʼ�], ���ް��� �հ�
    Taxinvoice.supplyCostTotal = "200000"
    
    '[�ʼ�] ���� �հ�
    Taxinvoice.taxTotal = "20000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + �����հ�
    Taxinvoice.totalAmount = "220000"
    
    '���� �� '�Ϸù�ȣ' �׸�
    Taxinvoice.serialNum = "123"
    
    '���� �� '��' �׸�, �ִ밪 32767
    Taxinvoice.kwon = "1"
    
    '���� �� 'ȣ' �׸�, �ִ밪 32767
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
    Taxinvoice.bankBookYN = False         '����纻 �̹��� ÷�ν� ����.
    

    '=========================================================================
    '         �������ݰ�꼭 ���� (�������ݰ�꼭 �ۼ��ÿ��� ����
    ' - �������ݰ�꼭 ���� ������ �����Ŵ��� �Ǵ� ���߰��̵� ��ũ ����
    ' - [����] �������ݰ�꼭 �ۼ���� �ȳ� - http://blog.linkhub.co.kr/650
    '========================================================================='
    
    ' ���������ڵ�, ���������� ���� 1~6�� ���ñ���
    Taxinvoice.modifyCode = ""
    
    ' �������ݰ�꼭�� ItemKey, ����Ȯ�� (GetInfo API)�� ������(ItemKey �׸�) Ȯ��
    Taxinvoice.originalTaxinvoiceKey = ""
        
    
    '=========================================================================
    '                            ���׸�(ǰ��) ����
    '=========================================================================
    
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20161010"   '�ŷ�����  yyyyMMdd
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
    newDetail2.purchaseDT = "20161010"   '�ŷ�����  yyyyMMdd
    newDetail2.itemName = "ǰ��2"        'ǰ���
    newDetail2.spec = "�԰�"             '�԰�
    newDetail2.qty = "1"                 '����
    newDetail2.unitCost = "100000"       '�ܰ�
    newDetail2.supplyCost = "100000"     '���ް���
    newDetail2.tax = "10000"             '����
    newDetail2.remark = "���"           '���
    
    Taxinvoice.detailList.Add newDetail2
    
    
    '=========================================================================
    '                              �߰������ ����
    ' - ���ݰ�꼭 ����ȳ� ������ ���Ź��� ���޹޴��� ����ڰ� �ټ��� ���
    ' ����� ������ �߰��Ͽ� ����ȳ������� �ټ����� ������ �� �ֽ��ϴ�.
    '=========================================================================
    
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                '�Ϸù�ȣ, 1���� ��������
    newContact.ContactName = "����� ����"  '����ڸ�
    newContact.email = "test2@test.com"     '����� �����ּ�
    Taxinvoice.addContactList.Add newContact
    
    Dim newContact2 As New PBTIContact
    newContact2.serialNum = 2                '�Ϸù�ȣ, 1���� ��������
    newContact2.ContactName = "����� ����"  '����ڸ�
    newContact2.email = "test2@test.com"     '����� �����ּ�
    
    Taxinvoice.addContactList.Add newContact2
    
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.Update(txtCorpNum.Text, KeyType, txtMgtKey.Text, Taxinvoice, False)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� ����� ������ �����մϴ�.
'=========================================================================

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�
    joinData.id = txtUserID.Text
    
    '����ڸ�
    joinData.personName = "����ڸ�_����"
    
    '����ó
    joinData.tel = "070-1234-1234"
    
    '�޴�����ȣ
    joinData.hp = "010-1234-1234"
    
    '�̸��� �ּ�
    joinData.email = "test@test.com"
    
    '�ѽ���ȣ
    joinData.fax = "070-1234-1234"
    
    '��ü��ȸ����, Ture-ȸ����ȸ, False-������
    joinData.searchAllAllowYN = True
    
    '������ ���ѿ���
    joinData.mgrYN = False
                
    Set Response = TaxinvoiceService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� �����մϴ�
'=========================================================================

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '��ǥ�ڸ�, �ִ� 30��
    CorpInfo.ceoname = "��ǥ��"
    
    '��ȣ, �ִ� 70��
    CorpInfo.corpName = "��ȣ"
    
    ' �ּ�, �ִ� 300��
    CorpInfo.addr = "����Ư����"
    
    '����, �ִ� 40��
    CorpInfo.bizType = "����"
    
    '����, �ִ� 40��
    CorpInfo.bizClass = "����"
    
    Set Response = TaxinvoiceService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "����޽��� : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

Private Sub Form_Load()
    ' ��� �ʱ�ȭ
    TaxinvoiceService.Initialize LinkID, SecretKey
    
    ' ����ȯ�� ������ True(���߿�), False(�����), ����� ��ȯ�� False�� ����.
    TaxinvoiceService.IsTest = True
        
    cboMgtKeyType.AddItem "SELL"
    cboMgtKeyType.AddItem "BUY"
    cboMgtKeyType.AddItem "TRUSTEE"
End Sub
