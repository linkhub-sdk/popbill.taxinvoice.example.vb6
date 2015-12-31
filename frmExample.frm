VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "팝빌 세금계산서 SDK 예제"
   ClientHeight    =   12000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14205
   LinkTopic       =   "Form1"
   ScaleHeight     =   12000
   ScaleWidth      =   14205
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton btnPopbillURL_CERT 
      Caption         =   " 공인인증서 등록 URL"
      Height          =   495
      Left            =   9360
      TabIndex        =   83
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton btnPopbillURL_CHRG 
      Caption         =   " 포인트 충전 URL"
      Height          =   495
      Left            =   9360
      TabIndex        =   82
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "문서 목록조회"
      Height          =   390
      Left            =   3080
      TabIndex        =   81
      Top             =   10820
      Width           =   1845
   End
   Begin VB.CommandButton btnUpdateCorpInfo 
      Caption         =   "회사정보 수정"
      Height          =   495
      Left            =   11640
      TabIndex        =   76
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton btnUpdateContact 
      Caption         =   "담당자 정보 수정"
      Height          =   495
      Left            =   7080
      TabIndex        =   74
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton btnListContact 
      Caption         =   "담당자 목록 조회"
      Height          =   495
      Left            =   7080
      TabIndex        =   73
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Frame Frame15 
      Caption         =   "회사정보 관련"
      Height          =   1695
      Left            =   11520
      TabIndex        =   71
      Top             =   960
      Width           =   2055
      Begin VB.CommandButton btnGetCorpInfo 
         Caption         =   "회사정보 조회"
         Height          =   495
         Left            =   120
         TabIndex        =   75
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton btnCheckID 
      Caption         =   "ID 중복 확인"
      Height          =   495
      Left            =   480
      TabIndex        =   69
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Frame Frame7 
      Caption         =   " 세금계산서 관련 기능"
      Height          =   8025
      Left            =   240
      TabIndex        =   14
      Top             =   3600
      Width           =   13575
      Begin VB.Frame Frame16 
         Caption         =   "정발행 (즉시발행) 프로세스"
         Height          =   3255
         Left            =   240
         TabIndex        =   77
         Top             =   840
         Width           =   3255
         Begin VB.CommandButton btnCancelIsse_2 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행취소"
            Height          =   495
            Left            =   240
            Style           =   1  '그래픽
            TabIndex        =   80
            Top             =   2110
            Width           =   975
         End
         Begin VB.CommandButton btnDelete_2 
            Caption         =   "삭제"
            Height          =   495
            Left            =   1920
            Style           =   1  '그래픽
            TabIndex        =   79
            Top             =   2110
            Width           =   975
         End
         Begin VB.CommandButton btnRegistIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "즉시발행"
            Height          =   495
            Left            =   480
            Style           =   1  '그래픽
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
            BackStyle       =   1  '투명하지 않음
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
         Caption         =   "유통메일목록"
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
         Caption         =   " 문서 정보 "
         Height          =   2760
         Left            =   7560
         TabIndex        =   62
         Top             =   5040
         Width           =   3210
         Begin VB.CommandButton btnGetEPrintUrl 
            Caption         =   "공급받는자 인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   68
            Top             =   1260
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPopUpURL 
            Caption         =   "문서 내용 보기 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   66
            Top             =   390
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPrintURL 
            Caption         =   "인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   65
            Top             =   825
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMassPrintURL 
            Caption         =   "다량 인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   64
            Top             =   1710
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMailURL 
            Caption         =   "이메일(공급받는자) 링크 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   63
            Top             =   2160
            Width           =   2745
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   " 기타 URL "
         Height          =   2295
         Left            =   11040
         TabIndex        =   57
         Top             =   5040
         Width           =   2265
         Begin VB.CommandButton btnGetURL_TBOX 
            Caption         =   "임시 문서함"
            Height          =   390
            Left            =   210
            TabIndex        =   61
            Top             =   390
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_SBOX 
            Caption         =   "매출 문서함"
            Height          =   390
            Left            =   210
            TabIndex        =   60
            Top             =   825
            Width           =   1845
         End
         Begin VB.CommandButton btn_GetURL_PBOX 
            Caption         =   "매입 문서함"
            Height          =   390
            Left            =   195
            TabIndex        =   59
            Top             =   1260
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_WRITE 
            Caption         =   "매출 작성"
            Height          =   390
            Left            =   195
            TabIndex        =   58
            Top             =   1710
            Width           =   1845
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   " 부가 서비스"
         Height          =   2775
         Left            =   5040
         TabIndex        =   53
         Top             =   5040
         Width           =   2265
         Begin VB.CommandButton btnDetachStatement 
            Caption         =   "전자명세서 첨부해제"
            Height          =   390
            Left            =   210
            TabIndex        =   85
            Top             =   2200
            Width           =   1845
         End
         Begin VB.CommandButton btnAttachStatement 
            Caption         =   "전자명세서 첨부"
            Height          =   390
            Left            =   210
            TabIndex        =   84
            Top             =   1750
            Width           =   1845
         End
         Begin VB.CommandButton btnSendEmail 
            Caption         =   "이메일 전송"
            Height          =   390
            Left            =   210
            TabIndex        =   56
            Top             =   390
            Width           =   1845
         End
         Begin VB.CommandButton btnSendSMS 
            Caption         =   "문자 전송"
            Height          =   390
            Left            =   210
            TabIndex        =   55
            Top             =   825
            Width           =   1845
         End
         Begin VB.CommandButton btnSendFAX 
            Caption         =   "팩스 전송"
            Height          =   390
            Left            =   210
            TabIndex        =   54
            Top             =   1290
            Width           =   1845
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   " 문서 정보 "
         Height          =   2775
         Left            =   2640
         TabIndex        =   48
         Top             =   5040
         Width           =   2265
         Begin VB.CommandButton btnGetDetailInfo 
            Caption         =   "문서 상세 정보"
            Height          =   390
            Left            =   195
            TabIndex        =   52
            Top             =   1710
            Width           =   1845
         End
         Begin VB.CommandButton btnGetLogs 
            Caption         =   "문서 이력"
            Height          =   390
            Left            =   195
            TabIndex        =   51
            Top             =   1260
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfos 
            Caption         =   "문서 정보(대량)"
            Height          =   390
            Left            =   210
            TabIndex        =   50
            Top             =   825
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfo 
            Caption         =   "문서 정보"
            Height          =   390
            Left            =   210
            TabIndex        =   49
            Top             =   390
            Width           =   1845
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   " 첨부파일 "
         Height          =   2280
         Left            =   240
         TabIndex        =   43
         Top             =   5055
         Width           =   2265
         Begin VB.CommandButton btnDeleteFile 
            Caption         =   "파일 삭제"
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
            Text            =   "파일아이디"
            Top             =   1245
            Width           =   1845
         End
         Begin VB.CommandButton btnGetFiles 
            Caption         =   "첨부 목록"
            Height          =   390
            Left            =   210
            TabIndex        =   45
            Top             =   795
            Width           =   1845
         End
         Begin VB.CommandButton btnAttachFile 
            Caption         =   "파일 첨부"
            Height          =   390
            Left            =   210
            TabIndex        =   44
            Top             =   345
            Width           =   1845
         End
      End
      Begin VB.CommandButton btnSendToNTS 
         BackColor       =   &H00C0C0FF&
         Caption         =   "국세청 즉시 전송"
         Height          =   495
         Left            =   4680
         Style           =   1  '그래픽
         TabIndex        =   42
         Top             =   4320
         Width           =   3000
      End
      Begin VB.Frame Frame9 
         Caption         =   " 역발행 세금계산서 프로세스 "
         Height          =   3345
         Left            =   9360
         TabIndex        =   22
         Top             =   840
         Width           =   3855
         Begin VB.CommandButton btnRefuse 
            BackColor       =   &H00C0C0FF&
            Caption         =   "거부"
            Height          =   375
            Left            =   2205
            Style           =   1  '그래픽
            TabIndex        =   41
            Top             =   1530
            Width           =   855
         End
         Begin VB.CommandButton btnRequestCancel 
            BackColor       =   &H00FFFFC0&
            Caption         =   "요청취소"
            Height          =   375
            Left            =   2205
            Style           =   1  '그래픽
            TabIndex        =   40
            Top             =   1050
            Width           =   855
         End
         Begin VB.CommandButton btnCancelIssue_rev 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행취소"
            Height          =   375
            Left            =   675
            Style           =   1  '그래픽
            TabIndex        =   39
            Top             =   2535
            Width           =   855
         End
         Begin VB.CommandButton btnDelete_rev 
            Caption         =   "삭제"
            Height          =   375
            Left            =   2670
            Style           =   1  '그래픽
            TabIndex        =   38
            Top             =   2550
            Width           =   855
         End
         Begin VB.CommandButton btnIssue_rev 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행"
            Height          =   360
            Left            =   735
            Style           =   1  '그래픽
            TabIndex        =   37
            Top             =   1980
            Width           =   720
         End
         Begin VB.CommandButton btnRequest 
            BackColor       =   &H00FFFFC0&
            Caption         =   "역)발행요청"
            Height          =   660
            Left            =   420
            Style           =   1  '그래픽
            TabIndex        =   36
            Top             =   1155
            Width           =   1350
         End
         Begin VB.CommandButton btnUpdate_rev 
            BackColor       =   &H00FFFFC0&
            Caption         =   "수정"
            Height          =   375
            Left            =   2475
            Style           =   1  '그래픽
            TabIndex        =   34
            Top             =   465
            Width           =   855
         End
         Begin VB.CommandButton btnRegister_rev 
            BackColor       =   &H00FFFFC0&
            Caption         =   "등록"
            Height          =   375
            Left            =   1515
            Style           =   1  '그래픽
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
            BackStyle       =   0  '투명
            Caption         =   "임시저장"
            Height          =   180
            Left            =   675
            TabIndex        =   35
            Top             =   540
            Width           =   720
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
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
         Caption         =   "정발행 (임시저장 발행, 발행예정) 프로세스"
         Height          =   3345
         Left            =   3720
         TabIndex        =   21
         Top             =   840
         Width           =   5415
         Begin VB.CommandButton btnCancelSend 
            BackColor       =   &H00C0C0FF&
            Caption         =   "취소"
            Height          =   375
            Left            =   3930
            Style           =   1  '그래픽
            TabIndex        =   32
            Top             =   1995
            Width           =   615
         End
         Begin VB.CommandButton btnDeny 
            BackColor       =   &H00FFFFC0&
            Caption         =   "거부"
            Height          =   375
            Left            =   3210
            Style           =   1  '그래픽
            TabIndex        =   31
            Top             =   1995
            Width           =   615
         End
         Begin VB.CommandButton btnAccept 
            BackColor       =   &H00FFFFC0&
            Caption         =   "승인"
            Height          =   375
            Left            =   2490
            Style           =   1  '그래픽
            TabIndex        =   30
            Top             =   1995
            Width           =   615
         End
         Begin VB.CommandButton btnSend 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행예정"
            Height          =   375
            Left            =   1650
            Style           =   1  '그래픽
            TabIndex        =   29
            Top             =   1425
            Width           =   855
         End
         Begin VB.CommandButton btnRegister 
            BackColor       =   &H00C0C0FF&
            Caption         =   "등록"
            Height          =   375
            Left            =   1305
            Style           =   1  '그래픽
            TabIndex        =   27
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton btnUpdate 
            BackColor       =   &H00C0C0FF&
            Caption         =   "수정"
            Height          =   375
            Left            =   2265
            Style           =   1  '그래픽
            TabIndex        =   26
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton btnDelete 
            Caption         =   "삭제"
            Height          =   375
            Left            =   3465
            Style           =   1  '그래픽
            TabIndex        =   25
            Top             =   2760
            Width           =   855
         End
         Begin VB.CommandButton btnCancelIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행취소"
            Height          =   375
            Left            =   360
            Style           =   1  '그래픽
            TabIndex        =   24
            Top             =   2730
            Width           =   855
         End
         Begin VB.CommandButton btnIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행"
            Height          =   525
            Left            =   345
            Style           =   1  '그래픽
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
            BackStyle       =   0  '투명
            Caption         =   "임시저장"
            Height          =   180
            Left            =   465
            TabIndex        =   28
            Top             =   555
            Width           =   720
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
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
            BackStyle       =   1  '투명하지 않음
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
         Caption         =   "관리번호 사용여부 확인"
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
         Caption         =   "문서관리번호( MgtKey) : "
         Height          =   180
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 팝빌 기본 API "
      Height          =   2895
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   13575
      Begin VB.Frame Frame6 
         Caption         =   " 공인인증서 관련"
         Height          =   1095
         Left            =   4440
         TabIndex        =   13
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnCertificateExpireDate 
            Caption         =   "인증서 만료일 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보"
         Height          =   2295
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "가입 여부 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "회원 가입"
            Height          =   495
            Left            =   120
            TabIndex        =   12
            Top             =   1560
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " 포인트 관련"
         Height          =   2295
         Left            =   1920
         TabIndex        =   8
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너 잔여포인트 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   70
            Top             =   1560
            Width           =   2175
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여포인트 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "요금 단가 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   9
            Top             =   960
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "담당자 관련"
         Height          =   2295
         Left            =   6720
         TabIndex        =   7
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "담당자 추가"
            Height          =   495
            Left            =   120
            TabIndex        =   72
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " 팝빌 기본 URL"
         Height          =   2295
         Left            =   9000
         TabIndex        =   5
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnGetPopbillURL_LOGIN 
            Caption         =   " 팝빌 로그인 URL"
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
      Caption         =   "팝빌회원 아이디 : "
      Height          =   180
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "팝빌회원 사업자번호 :"
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

'링크아이디
Private Const LinkID = "TESTER"
'비밀키. 유출에 주의하시기 바랍니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.Accept(txtCorpNum.Text, KeyType, txtMgtKey.Text, "발행예정 승인 메모", txtUserID.Text)
    
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
            MsgBox "관리번호 형태를 선택해주세요."
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    SubItemCode = 121           '첨부할 전자명세서 종류코드, 121-거래명세서, 122-청구서, 123-견적서, 124-발주서, 125-입금표,126-영수증
    SubMgtKey = "20151223-01"   '첨부할 전자명세서 관리번호
        
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, "발행 취소 메모", txtUserID.Text)
    
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, "발행 취소 메모", txtUserID.Text)
    
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, "발행 취소 메모", txtUserID.Text)
    
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.CancelSend(txtCorpNum.Text, KeyType, txtMgtKey.Text, "발행예정 취소 메모", txtUserID.Text)
    
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
    
    MsgBox "인증서만료일 : " + expireDate
 
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
            MsgBox "관리번호 형태를 선택해주세요."
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
            MsgBox "관리번호 형태를 선택해주세요."
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
            MsgBox "관리번호 형태를 선택해주세요."
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
            MsgBox "관리번호 형태를 선택해주세요."
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
            MsgBox "관리번호 형태를 선택해주세요."
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.Deny(txtCorpNum.Text, KeyType, txtMgtKey.Text, "발행예정 거부 메모", txtUserID.Text)
    
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    SubItemCode = 121           '첨부할 전자명세서 종류코드, 121-거래명세서, 122-청구서, 123-견적서, 124-발주서, 125-입금표,126-영수증
    SubMgtKey = "20151223-01"   '첨부할 전자명세서 관리번호
        
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
    
    MsgBox "잔여포인트 : " + CStr(balance)
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
            MsgBox "관리번호 형태를 선택해주세요."
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

    '''  상세내역 생략 '''
    
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
            MsgBox "관리번호 형태를 선택해주세요."
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
            MsgBox "관리번호 형태를 선택해주세요."
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
    
    '관리번호 배열, 최대 1000건
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
            MsgBox "관리번호 형태를 선택해주세요."
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
    
    MsgBox "잔여포인트 : " + CStr(balance)
    
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
            MsgBox "관리번호 형태를 선택해주세요."
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
            MsgBox "관리번호 형태를 선택해주세요."
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.Issue(txtCorpNum.Text, KeyType, txtMgtKey.Text, "발행메모", "", True, txtUserID.Text)
    
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.Issue(txtCorpNum.Text, KeyType, txtMgtKey.Text, "발행메모", "", False, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    joinData.LinkID = LinkID '링크 아이디
    joinData.CorpNum = "1231212312" '사업자번호 "-" 제외.
    joinData.ceoname = "대표자성명"
    joinData.corpName = "회원상호"
    joinData.addr = "주소"
    joinData.bizType = "업태"
    joinData.bizClass = "업종"
    joinData.id = "userid"      '6자 이상 20자 미만.
    joinData.pwd = "pwd_must_be_long_enough"    '6자 이상 20자 미만.
    joinData.ContactName = "담당자성명"
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.Refuse(txtCorpNum.Text, KeyType, txtMgtKey.Text, "역)발행 요청 거부 메모", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    joinData.id = "testkorea_20151007"      '담당자 아이디
    joinData.pwd = "test@test.com"          '비밀번호
    joinData.personName = "담당자명"        '담당자명
    joinData.tel = "070-1234-1234"          '연락처
    joinData.hp = "010-1234-1234"           '휴대폰번호
    joinData.email = "test@test.com"        '이메일 주소
    joinData.fax = "070-1234-1234"          '팩스번호
    joinData.searchAllAllowYN = True        '전체조회여부, Ture-회사조회, False-개인조회
    joinData.mgrYN = False                  '관리자 권한여부
        
    Set Response = TaxinvoiceService.RegistContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
    
End Sub

Private Sub btnRegister_Click()
    Dim Taxinvoice As New PBTaxinvoice
    
    Taxinvoice.writeDate = "20151012"             '필수, 기재상 작성일자
    Taxinvoice.chargeDirection = "정과금"         '필수, {정과금, 역과금}
    Taxinvoice.issueType = "정발행"               '필수, {정발행, 역발행, 위수탁}
    Taxinvoice.purposeType = "영수"               '필수, {영수, 청구}
    Taxinvoice.issueTiming = "직접발행"           '필수, {직접발행, 승인시자동발행}
    Taxinvoice.taxType = "과세"                   '필수, {과세, 영세, 면세}
        
    Taxinvoice.invoicerCorpNum = "1234567890"     '공급자 사업자번호
    Taxinvoice.invoicerTaxRegID = ""              '종사업자 식별번호. 필요시 기재. 형식은 숫자 4자리.
    Taxinvoice.invoicerCorpName = "공급자 상호"
    Taxinvoice.invoicerMgtKey = txtMgtKey.Text    '공급자 문서관리번호, 1~24자리, 숫자,영문,'-','_' 조합하여 임의로 구성
    Taxinvoice.invoicerCEOName = "공급자"" 대표자 성명"
    Taxinvoice.invoicerAddr = "공급자 주소"
    Taxinvoice.invoicerBizClass = "공급자 업종"
    Taxinvoice.invoicerBizType = "공급자 업태,업태2"
    Taxinvoice.invoicerContactName = "공급자 담당자명"
    Taxinvoice.invoicerEmail = "test@test.com"
    Taxinvoice.invoicerTEL = "070-7070-0707"
    Taxinvoice.invoicerHP = "010-000-2222"
    Taxinvoice.invoicerSMSSendYN = False            '정발행시(공급자->공급받는자) 문자발송여부
    
    Taxinvoice.invoiceeType = "사업자"             '공급받는자 구분, {사업자, 개인, 외국인} 중 기재
    Taxinvoice.invoiceeCorpNum = "8888888888"      '공급받는자 사업자번호
    Taxinvoice.invoiceeCorpName = "공급받는자 상호"
    Taxinvoice.invoiceeMgtKey = ""                 '공급받는자 문서관리번호(역발행시에만 필수)
    Taxinvoice.invoiceeCEOName = "공급받는자 대표자 성명"
    Taxinvoice.invoiceeAddr = "공급받는자 주소"
    Taxinvoice.invoiceeBizClass = "공급받는자 업종"
    Taxinvoice.invoiceeBizType = "공급받는자 업태"
    Taxinvoice.invoiceeContactName1 = "공급받는자 담당자명"
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    Taxinvoice.invoiceeHP1 = "010-111-222"
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    Taxinvoice.invoiceeSMSSendYN = False          '역발행시(공급받는자->공급자) 문자발송여부
            
    Taxinvoice.supplyCostTotal = "100000"         '필수 공급가액 합계
    Taxinvoice.taxTotal = "10000"                 '필수 세액 합계
    Taxinvoice.totalAmount = "110000"             '필수 합계금액.  공급가액 + 세액
    
    Taxinvoice.modifyCode = "" '수정세금계산서 작성시 1~6까지 선택기재.
    Taxinvoice.originalTaxinvoiceKey = "" '수정세금계산서 작성시 원본세금계산서의 ItemKey기재. ItemKey는 문서확인.
    Taxinvoice.serialNum = "123"  '일련번호
    Taxinvoice.cash = ""          '현금
    Taxinvoice.chkBill = ""       '수표
    Taxinvoice.note = ""          '어음
    Taxinvoice.credit = ""        '외상미수금
    Taxinvoice.remark1 = "비고1"
    Taxinvoice.remark2 = "비고2"
    Taxinvoice.remark3 = "비고3"
    Taxinvoice.kwon = "1"           '권
    Taxinvoice.ho = "1"             '호
    
    Taxinvoice.businessLicenseYN = False '사업자등록증 이미지 첨부시 설정.
    Taxinvoice.bankBookYN = False         '통장사본 이미지 첨부시 설정.
    
    
    '상세항목 추가.
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20140410"   '거래일자  yyyyMMdd
    newDetail.itemName = "품명"
    newDetail.spec = "규격"
    newDetail.qty = "1" '수량           ' 소숫점 2자리까지 문자열로 기재가능
    newDetail.unitCost = "100000"       ' 단가, 소숫점 2자리까지 문자열로 기재가능
    newDetail.supplyCost = "100000"     ' 공급가액
    newDetail.tax = "10000"             ' 세액
    newDetail.remark = "비고"           ' 비고
    
    Taxinvoice.detailList.Add newDetail
    
    Set newDetail = New PBTIDetail
    newDetail.serialNum = 2
    newDetail.itemName = "품명2"
    
    Taxinvoice.detailList.Add newDetail
    
    
    '추가담당자 추가. 옵션.
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                '일련번호, 1부터 순차기재
    newContact.ContactName = "담당자 성명"
    newContact.email = "test2@test.com"
    
    Taxinvoice.addContactList.Add newContact
        
    
    Dim Response As PBResponse
    
    'Register(사업자번호, 세금계산서 객체, 거래명세서 동시작성여부, 팝빌회워아이디)
    Set Response = TaxinvoiceService.Register(txtCorpNum.Text, Taxinvoice, False, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
    

End Sub

Private Sub btnRegister_rev_Click()
    Dim Taxinvoice As New PBTaxinvoice
    
    Taxinvoice.writeDate = "20151008"             '필수, 기재상 작성일자
    Taxinvoice.chargeDirection = "정과금"         '필수, {정과금, 역과금}
    Taxinvoice.issueType = "역발행"               '필수, {정발행, 역발행, 위수탁}
    Taxinvoice.purposeType = "영수"               '필수, {영수, 청구}
    Taxinvoice.issueTiming = "직접발행"           '필수, {직접발행, 승인시자동발행}
    Taxinvoice.taxType = "과세"                   '필수, {과세, 영세, 면세}
    
    
    Taxinvoice.invoicerCorpNum = "8888888888"
    Taxinvoice.invoicerTaxRegID = "" '종사업자 식별번호. 필요시 기재. 형식은 숫자 4자리.
    Taxinvoice.invoicerCorpName = "공급자 상호"
    Taxinvoice.invoicerMgtKey = ""
    Taxinvoice.invoicerCEOName = "공급자"" 대표자 성명"
    Taxinvoice.invoicerAddr = "공급자 주소"
    Taxinvoice.invoicerBizClass = "공급자 업종"
    Taxinvoice.invoicerBizType = "공급자 업태,업태2"
    Taxinvoice.invoicerContactName = "공급자 담당자명"
    Taxinvoice.invoicerEmail = "test@test.com"
    Taxinvoice.invoicerTEL = "070-7070-0707"
    Taxinvoice.invoicerHP = "010-000-2222"
    Taxinvoice.invoicerSMSSendYN = True '발행시 문자발송기능 사용시 활용
    
    Taxinvoice.invoiceeType = "사업자"
    Taxinvoice.invoiceeCorpNum = "1231212312"
    Taxinvoice.invoiceeCorpName = "공급받는자 상호"
    Taxinvoice.invoiceeMgtKey = txtMgtKey.Text
    Taxinvoice.invoiceeCEOName = "공급받는자 대표자 성명"
    Taxinvoice.invoiceeAddr = "공급받는자 주소"
    Taxinvoice.invoiceeBizClass = "공급받는자 업종"
    Taxinvoice.invoiceeBizType = "공급받는자 업태"
    Taxinvoice.invoiceeContactName1 = "공급받는자 담당자명"
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    Taxinvoice.supplyCostTotal = "100000"         '필수 공급가액 합계
    Taxinvoice.taxTotal = "10000"                 '필수 세액 합계
    Taxinvoice.totalAmount = "110000"             '필수 합계금액.  공급가액 + 세액
    
    Taxinvoice.modifyCode = "" '수정세금계산서 작성시 1~6까지 선택기재.
    Taxinvoice.originalTaxinvoiceKey = "" '수정세금계산서 작성시 원본세금계산서의 ItemKey기재. ItemKey는 문서확인.
    Taxinvoice.serialNum = "123"
    Taxinvoice.cash = ""          '현금
    Taxinvoice.chkBill = ""       '수표
    Taxinvoice.note = ""          '어음
    Taxinvoice.credit = ""        '외상미수금
    Taxinvoice.remark1 = "비고1"
    Taxinvoice.remark2 = "비고2"
    Taxinvoice.remark3 = "비고3"
    Taxinvoice.kwon = "1"
    Taxinvoice.ho = "1"
    
    Taxinvoice.businessLicenseYN = False '사업자등록증 이미지 첨부시 설정.
    Taxinvoice.bankBookYN = False         '통장사본 이미지 첨부시 설정.
       
    
    '상세항목 추가.
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1
    newDetail.purchaseDT = "20140410"
    newDetail.itemName = "품명"
    newDetail.spec = "규격"
    newDetail.qty = "1" '수량
    newDetail.unitCost = "100000"
    newDetail.supplyCost = "100000"
    newDetail.tax = "10000"
    newDetail.remark = "비고"
    
    Taxinvoice.detailList.Add newDetail
    
    Set newDetail = New PBTIDetail
    newDetail.serialNum = 2
    newDetail.itemName = "품명2"
    
    Taxinvoice.detailList.Add newDetail
    
    
    '추가담당자 추가. 옵션.
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.ContactName = "담당자 성명"
    newContact.email = "test2@test.com"
    
    Taxinvoice.addContactList.Add newContact
    
    
    Dim Response As PBResponse
    'Register(사업자번호, 세금계산서 객체, 거래명세서 동시작성여부, 팝빌회원아이디)
    Set Response = TaxinvoiceService.Register(txtCorpNum.Text, Taxinvoice, False, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
End Sub

Private Sub btnRegistIssue_Click()
    
    Dim Taxinvoice As New PBTaxinvoice
    
    Taxinvoice.writeSpecification = False         '거래명세서 동시작성 여부
    Taxinvoice.forceIssue = False                 '지연발행 강제여부
    Taxinvoice.memo = ""                          '메모
    Taxinvoice.emailSubject = ""                  '안내메일 제목, 공백처리시 기본제목으로 전송
    Taxinvoice.dealInvoiceMgtKey = ""             '거래명세서 동시작성시 명세서 관리번호, 미기재시 세금계산서 관리번호로 자동작성
        
    Taxinvoice.writeDate = "20151012"             '필수, 기재상 작성일자
    Taxinvoice.chargeDirection = "정과금"         '필수, {정과금, 역과금}
    Taxinvoice.issueType = "정발행"               '필수, {정발행, 역발행, 위수탁}
    Taxinvoice.purposeType = "영수"               '필수, {영수, 청구}
    Taxinvoice.issueTiming = "직접발행"           '필수, {직접발행, 승인시자동발행}
    Taxinvoice.taxType = "과세"                   '필수, {과세, 영세, 면세}
        
    Taxinvoice.invoicerCorpNum = "1234567890"     '공급자 사업자번호
    Taxinvoice.invoicerTaxRegID = ""              '종사업자 식별번호. 필요시 기재. 형식은 숫자 4자리.
    Taxinvoice.invoicerCorpName = "공급자 상호"
    Taxinvoice.invoicerMgtKey = txtMgtKey.Text    '공급자 문서관리번호, 1~24자리, 숫자,영문,'-','_' 조합하여 임의로 구성
    Taxinvoice.invoicerCEOName = "공급자"" 대표자 성명"
    Taxinvoice.invoicerAddr = "공급자 주소"
    Taxinvoice.invoicerBizClass = "공급자 업종"
    Taxinvoice.invoicerBizType = "공급자 업태,업태2"
    Taxinvoice.invoicerContactName = "공급자 담당자명"
    Taxinvoice.invoicerEmail = "test@test.com"
    Taxinvoice.invoicerTEL = "070-7070-0707"
    Taxinvoice.invoicerHP = "010-000-2222"
    Taxinvoice.invoicerSMSSendYN = False            '정발행시(공급자->공급받는자) 문자발송여부
    
    Taxinvoice.invoiceeType = "사업자"             '공급받는자 구분, {사업자, 개인, 외국인} 중 기재
    Taxinvoice.invoiceeCorpNum = "8888888888"      '공급받는자 사업자번호
    Taxinvoice.invoiceeCorpName = "공급받는자 상호"
    Taxinvoice.invoiceeMgtKey = ""                 '공급받는자 문서관리번호(역발행시에만 필수)
    Taxinvoice.invoiceeCEOName = "공급받는자 대표자 성명"
    Taxinvoice.invoiceeAddr = "공급받는자 주소"
    Taxinvoice.invoiceeBizClass = "공급받는자 업종"
    Taxinvoice.invoiceeBizType = "공급받는자 업태"
    Taxinvoice.invoiceeContactName1 = "공급받는자 담당자명"
    Taxinvoice.invoiceeEmail1 = "test@test.com"
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    Taxinvoice.invoiceeHP1 = "010-111-222"
    Taxinvoice.invoiceeSMSSendYN = False          '역발행시(공급받는자->공급자) 문자발송여부
            
    Taxinvoice.supplyCostTotal = "100000"         '필수 공급가액 합계
    Taxinvoice.taxTotal = "10000"                 '필수 세액 합계
    Taxinvoice.totalAmount = "110000"             '필수 합계금액.  공급가액 + 세액
    
    Taxinvoice.modifyCode = "" '수정세금계산서 작성시 1~6까지 선택기재.
    Taxinvoice.originalTaxinvoiceKey = "" '수정세금계산서 작성시 원본세금계산서의 ItemKey기재. ItemKey는 문서확인.
    Taxinvoice.serialNum = "123"  '일련번호
    Taxinvoice.cash = ""          '현금
    Taxinvoice.chkBill = ""       '수표
    Taxinvoice.note = ""          '어음
    Taxinvoice.credit = ""        '외상미수금
    Taxinvoice.remark1 = "비고1"
    Taxinvoice.remark2 = "비고2"
    Taxinvoice.remark3 = "비고3"
    Taxinvoice.kwon = "1"           '권
    Taxinvoice.ho = "1"             '호
    
    Taxinvoice.businessLicenseYN = False '사업자등록증 이미지 첨부시 설정.
    Taxinvoice.bankBookYN = False         '통장사본 이미지 첨부시 설정.
    
    
    '상세항목 추가.
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20140410"   '거래일자  yyyyMMdd
    newDetail.itemName = "품명"
    newDetail.spec = "규격"
    newDetail.qty = "1" '수량           ' 소숫점 2자리까지 문자열로 기재가능
    newDetail.unitCost = "100000"       '단가 소숫점 2자리까지 문자열로 기재가능
    newDetail.supplyCost = "100000"     '공급가액
    newDetail.tax = "10000"             '세액
    newDetail.remark = "비고"           '비고
    
    Taxinvoice.detailList.Add newDetail
    
    Set newDetail = New PBTIDetail
    newDetail.serialNum = 2
    newDetail.itemName = "품명2"
    
    Taxinvoice.detailList.Add newDetail
    
    
    '추가담당자 추가. 옵션.
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                '일련번호 1부터 순차기재
    newContact.ContactName = "담당자 성명"  '담당자명
    newContact.email = "test2@test.com"     '이메일주소
    
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.Request(txtCorpNum.Text, KeyType, txtMgtKey.Text, "역)발행 요청 메모", txtUserID.Text)
    
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.CancelRequest(txtCorpNum.Text, KeyType, txtMgtKey.Text, "역)발행 요청 취소 메모", txtUserID.Text)
    
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    
    DType = "I"             '[필수] 일자유형, R-등록일시 W-작성일자 I-발행일시 중 택1
    SDate = "20151206"      '[필수] 시작일자, yyyyMMdd
    EDate = "20151231"      '[필수] 종료일자, yyyyMMdd
    
    State.Add "100"         '전송상태값 배열, 미기재시 전체상태조회, 문서상태값 3자리숫자 작성
    State.Add "2**"         '2,3번째 와일드카드 가능
    State.Add "3**"
    
    TType.Add "N"           '문서유형 배열, N-일반 M-수정 중 선택, 미기재시 전체조회
    TType.Add "M"
    
    taxType.Add "T"         '과세형태 배열, T-과세, N-면세 Z-영세 중 선택, 미기재시 전체조회
    taxType.Add "N"
    taxType.Add "Z"
    
    LateOnly = ""           '지연발행 여부, 0-정상발행분만 조회 1-지연발행분만조회, 공백처리시 전체조회
    
    Page = 1                '페이지 번호
    PerPage = 10            '페이지 목록개수, 최대 1000건
    
    Order = "A"             '정렬방향, D-내림차순(기본값), A-오름차순
    
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    emailSubject = ""       '발행예정 안내메일 제목, 공백처리시 기본제목으로 전송
    memo = "발행예정 메모"  '메모
    
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    receiveEmail = "test@test.com" '수신자 메일주소
    
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
        
    senderNum = "07075103710"     '발신번호
    receiveNum = "111-222-4444"   '수신번호
        
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    senderNum = "07075103710"
    receiveNum = "111-2222-4444"
    Contents = "문자 내용, 90Byte초과시 길이가 조정되어 전송됨"
    
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
            MsgBox "관리번호 형태를 선택해주세요."
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
    
    MsgBox "발행단가 : " + CStr(unitCost)
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Dim Taxinvoice As New PBTaxinvoice
    
    Taxinvoice.writeDate = "20140319"             '필수, 기재상 작성일자
    Taxinvoice.chargeDirection = "정과금"         '필수, {정과금, 역과금}
    Taxinvoice.issueType = "정발행"               '필수, {정발행, 역발행, 위수탁}
    Taxinvoice.purposeType = "영수"               '필수, {영수, 청구}
    Taxinvoice.issueTiming = "직접발행"           '필수, {직접발행, 승인시자동발행}
    Taxinvoice.taxType = "과세"                   '필수, {과세, 영세, 면세}
    
    
    Taxinvoice.invoicerCorpNum = "1231212312"
    Taxinvoice.invoicerTaxRegID = "" '종사업자 식별번호. 필요시 기재. 형식은 숫자 4자리.
    Taxinvoice.invoicerCorpName = "공급자 상호"
    Taxinvoice.invoicerMgtKey = txtMgtKey.Text
    Taxinvoice.invoicerCEOName = "공급자"" 대표자 성명"
    Taxinvoice.invoicerAddr = "공급자 주소"
    Taxinvoice.invoicerBizClass = "공급자 업종"
    Taxinvoice.invoicerBizType = "공급자 업태,업태2"
    Taxinvoice.invoicerContactName = "공급자 담당자명"
    Taxinvoice.invoicerEmail = "test@test.com"
    Taxinvoice.invoicerTEL = "070-7070-0707"
    Taxinvoice.invoicerHP = "010-000-2222"
    Taxinvoice.invoicerSMSSendYN = True '발행시 문자발송기능 사용시 활용
    
    Taxinvoice.invoiceeType = "사업자"
    Taxinvoice.invoiceeCorpNum = "8888888888"
    Taxinvoice.invoiceeCorpName = "공급받는자 상호"
    Taxinvoice.invoiceeMgtKey = ""
    Taxinvoice.invoiceeCEOName = "공급받는자 대표자 성명"
    Taxinvoice.invoiceeAddr = "공급받는자 주소"
    Taxinvoice.invoiceeBizClass = "공급받는자 업종"
    Taxinvoice.invoiceeBizType = "공급받는자 업태"
    Taxinvoice.invoiceeContactName1 = "공급받는자 담당자명"
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    Taxinvoice.supplyCostTotal = "100000"         '필수 공급가액 합계
    Taxinvoice.taxTotal = "10000"                 '필수 세액 합계
    Taxinvoice.totalAmount = "110000"             '필수 합계금액.  공급가액 + 세액
    
    Taxinvoice.modifyCode = "" '수정세금계산서 작성시 1~6까지 선택기재.
    Taxinvoice.originalTaxinvoiceKey = "" '수정세금계산서 작성시 원본세금계산서의 ItemKey기재. ItemKey는 문서확인.
    Taxinvoice.serialNum = "123"
    Taxinvoice.cash = ""          '현금
    Taxinvoice.chkBill = ""       '수표
    Taxinvoice.note = ""          '어음
    Taxinvoice.credit = ""        '외상미수금
    Taxinvoice.remark1 = "비고1"
    Taxinvoice.remark2 = "비고2"
    Taxinvoice.remark3 = "비고3"
    Taxinvoice.kwon = "1"
    Taxinvoice.ho = "1"
    
    Taxinvoice.businessLicenseYN = False '사업자등록증 이미지 첨부시 설정.
    Taxinvoice.bankBookYN = False         '통장사본 이미지 첨부시 설정.
        
    
    '상세항목 추가.
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1
    newDetail.purchaseDT = "20140410"
    newDetail.itemName = "품명"
    newDetail.spec = "규격"
    newDetail.qty = "1" '수량
    newDetail.unitCost = "100000"
    newDetail.supplyCost = "100000"
    newDetail.tax = "10000"
    newDetail.remark = "비고"
    
    Taxinvoice.detailList.Add newDetail
    
    Set newDetail = New PBTIDetail
    newDetail.serialNum = 2
    newDetail.itemName = "품명2_수정됨"
    
    Taxinvoice.detailList.Add newDetail
    
    
    '추가담당자 추가. 옵션.
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.ContactName = "담당자 성명"
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
    
    Taxinvoice.writeDate = "20140319"             '필수, 기재상 작성일자
    Taxinvoice.chargeDirection = "정과금"         '필수, {정과금, 역과금}
    Taxinvoice.issueType = "역발행"               '필수, {정발행, 역발행, 위수탁}
    Taxinvoice.purposeType = "영수"               '필수, {영수, 청구}
    Taxinvoice.issueTiming = "직접발행"           '필수, {직접발행, 승인시자동발행}
    Taxinvoice.taxType = "과세"                   '필수, {과세, 영세, 면세}
    
    
    Taxinvoice.invoicerCorpNum = "8888888888"
    Taxinvoice.invoicerTaxRegID = "" '종사업자 식별번호. 필요시 기재. 형식은 숫자 4자리.
    Taxinvoice.invoicerCorpName = "공급자 상호"
    Taxinvoice.invoicerMgtKey = ""
    Taxinvoice.invoicerCEOName = "공급자"" 대표자 성명"
    Taxinvoice.invoicerAddr = "공급자 주소"
    Taxinvoice.invoicerBizClass = "공급자 업종"
    Taxinvoice.invoicerBizType = "공급자 업태,업태2"
    Taxinvoice.invoicerContactName = "공급자 담당자명"
    Taxinvoice.invoicerEmail = "test@test.com"
    Taxinvoice.invoicerTEL = "070-7070-0707"
    Taxinvoice.invoicerHP = "010-000-2222"
    Taxinvoice.invoicerSMSSendYN = True '발행시 문자발송기능 사용시 활용
    
    Taxinvoice.invoiceeType = "사업자"
    Taxinvoice.invoiceeCorpNum = "1231212312"
    Taxinvoice.invoiceeCorpName = "공급받는자 상호"
    Taxinvoice.invoiceeMgtKey = txtMgtKey.Text
    Taxinvoice.invoiceeCEOName = "공급받는자 대표자 성명"
    Taxinvoice.invoiceeAddr = "공급받는자 주소"
    Taxinvoice.invoiceeBizClass = "공급받는자 업종"
    Taxinvoice.invoiceeBizType = "공급받는자 업태"
    Taxinvoice.invoiceeContactName1 = "공급받는자 담당자명"
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    Taxinvoice.supplyCostTotal = "100000"         '필수 공급가액 합계
    Taxinvoice.taxTotal = "10000"                 '필수 세액 합계
    Taxinvoice.totalAmount = "110000"             '필수 합계금액.  공급가액 + 세액
    
    Taxinvoice.modifyCode = "" '수정세금계산서 작성시 1~6까지 선택기재.
    Taxinvoice.originalTaxinvoiceKey = "" '수정세금계산서 작성시 원본세금계산서의 ItemKey기재. ItemKey는 문서확인.
    Taxinvoice.serialNum = "123"
    Taxinvoice.cash = ""          '현금
    Taxinvoice.chkBill = ""       '수표
    Taxinvoice.note = ""          '어음
    Taxinvoice.credit = ""        '외상미수금
    Taxinvoice.remark1 = "비고1"
    Taxinvoice.remark2 = "비고2"
    Taxinvoice.remark3 = "비고3"
    Taxinvoice.kwon = "1"
    Taxinvoice.ho = "1"
    
    Taxinvoice.businessLicenseYN = False '사업자등록증 이미지 첨부시 설정.
    Taxinvoice.bankBookYN = False         '통장사본 이미지 첨부시 설정.
    
    '상세항목 추가.
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1
    newDetail.purchaseDT = "20140410"
    newDetail.itemName = "품명"
    newDetail.spec = "규격"
    newDetail.qty = "1" '수량
    newDetail.unitCost = "100000"
    newDetail.supplyCost = "100000"
    newDetail.tax = "10000"
    newDetail.remark = "비고"
    
    Taxinvoice.detailList.Add newDetail
    
    Set newDetail = New PBTIDetail
    newDetail.serialNum = 2
    newDetail.itemName = "품명2_수정됨"
    
    Taxinvoice.detailList.Add newDetail
    
    
    '추가담당자 추가. 옵션.
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.ContactName = "담당자 성명"
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
    
    joinData.personName = "담당자명_수정"  '담당자명
    joinData.tel = "070-1234-1234"         '연락처
    joinData.hp = "010-1234-1234"          '휴대폰번호
    joinData.email = "test@test.com"       '이메일 주소
    joinData.fax = "070-1234-1234"         '팩스번호
    joinData.searchAllAllowYN = True       '전체조회여부, Ture-회사조회, False-개인조
    joinData.mgrYN = False                 '관리자 권한여부
                
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
    
    CorpInfo.ceoname = "대표자"         '대표자명
    CorpInfo.corpName = "상호"          '상호명
    CorpInfo.addr = "서울특별시"        '주소
    CorpInfo.bizType = "업태"           '업태
    CorpInfo.bizClass = "업종"          '업종
    
    Set Response = TaxinvoiceService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(TaxinvoiceService.LastErrCode) + "] " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub


Private Sub Form_Load()
    '모듈 초기화
    TaxinvoiceService.Initialize LinkID, SecretKey
    
    '연동환경 설정값 True(테스트용), False(상업용)
    TaxinvoiceService.IsTest = True
        
    cboMgtKeyType.AddItem "SELL"
    cboMgtKeyType.AddItem "BUY"
    cboMgtKeyType.AddItem "TRUSTEE"
End Sub

