VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "팝빌 세금계산서 SDK 예제"
   ClientHeight    =   12000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14490
   LinkTopic       =   "Form1"
   ScaleHeight     =   12000
   ScaleWidth      =   14490
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton btnPopbillURL_CHRG 
      Caption         =   " 포인트 충전 URL"
      Height          =   410
      Left            =   9360
      TabIndex        =   82
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "문서 목록조회"
      Height          =   390
      Left            =   3080
      TabIndex        =   81
      Top             =   10920
      Width           =   1845
   End
   Begin VB.CommandButton btnUpdateCorpInfo 
      Caption         =   "회사정보 수정"
      Height          =   410
      Left            =   12120
      TabIndex        =   76
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton btnUpdateContact 
      Caption         =   "담당자 정보 수정"
      Height          =   410
      Left            =   7080
      TabIndex        =   74
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton btnListContact 
      Caption         =   "담당자 목록 조회"
      Height          =   410
      Left            =   7080
      TabIndex        =   73
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Frame Frame15 
      Caption         =   "회사정보 관련"
      Height          =   2415
      Left            =   12000
      TabIndex        =   71
      Top             =   960
      Width           =   2055
      Begin VB.CommandButton btnGetCorpInfo 
         Caption         =   "회사정보 조회"
         Height          =   410
         Left            =   120
         TabIndex        =   75
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton btnCheckID 
      Caption         =   "ID 중복 확인"
      Height          =   410
      Left            =   480
      TabIndex        =   69
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame Frame7 
      Caption         =   " 세금계산서 관련 기능"
      Height          =   8025
      Left            =   240
      TabIndex        =   14
      Top             =   3720
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
            TabIndex        =   84
            Top             =   2200
            Width           =   1845
         End
         Begin VB.CommandButton btnAttachStatement 
            Caption         =   "전자명세서 첨부"
            Height          =   390
            Left            =   210
            TabIndex        =   83
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
         Left            =   4560
         Style           =   1  '그래픽
         TabIndex        =   42
         Top             =   4320
         Width           =   4440
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
            Left            =   2160
            Style           =   1  '그래픽
            TabIndex        =   41
            Top             =   1530
            Width           =   855
         End
         Begin VB.CommandButton btnRequestCancel 
            BackColor       =   &H00FFFFC0&
            Caption         =   "요청취소"
            Height          =   375
            Left            =   2160
            Style           =   1  '그래픽
            TabIndex        =   40
            Top             =   1050
            Width           =   855
         End
         Begin VB.CommandButton btnCancelIssue_rev 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행취소"
            Height          =   495
            Left            =   600
            Style           =   1  '그래픽
            TabIndex        =   39
            Top             =   2520
            Width           =   1095
         End
         Begin VB.CommandButton btnDelete_rev 
            Caption         =   "삭제"
            Height          =   375
            Left            =   2670
            Style           =   1  '그래픽
            TabIndex        =   38
            Top             =   2550
            Width           =   975
         End
         Begin VB.CommandButton btnIssue_rev 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행"
            Height          =   480
            Left            =   600
            Style           =   1  '그래픽
            TabIndex        =   37
            Top             =   1920
            Width           =   1080
         End
         Begin VB.CommandButton btnRequest 
            BackColor       =   &H00FFFFC0&
            Caption         =   "역)발행요청"
            Height          =   660
            Left            =   540
            Style           =   1  '그래픽
            TabIndex        =   36
            Top             =   1155
            Width           =   1110
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
            Width           =   975
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
      Width           =   14055
      Begin VB.Frame Frame6 
         Caption         =   " 공인인증서 관련"
         Height          =   2415
         Left            =   4440
         TabIndex        =   13
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnPopbillURL_CERT 
            Caption         =   " 공인인증서 등록 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   86
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton btnCertificateExpireDate 
            Caption         =   "인증서 만료일 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보"
         Height          =   2415
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "가입 여부 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "회원 가입"
            Height          =   410
            Left            =   120
            TabIndex        =   12
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " 포인트 관련"
         Height          =   2415
         Left            =   1920
         TabIndex        =   8
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "과금정보 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   85
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너 잔여포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   70
            Top             =   1320
            Width           =   2175
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   2175
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "요금 단가 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   9
            Top             =   1800
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "담당자 관련"
         Height          =   2415
         Left            =   6720
         TabIndex        =   7
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "담당자 추가"
            Height          =   410
            Left            =   120
            TabIndex        =   72
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " 팝빌 기본 URL"
         Height          =   2415
         Left            =   9000
         TabIndex        =   5
         Top             =   360
         Width           =   2655
         Begin VB.CommandButton btnGetPopbillURL_SEAL 
            Caption         =   "인감 및 첨부문서 등록 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   87
            Top             =   1320
            Width           =   2415
         End
         Begin VB.CommandButton btnGetPopbillURL_LOGIN 
            Caption         =   " 팝빌 로그인 URL"
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
'=========================================================================
'
' 팝빌 전자세금계산서 API VB 6.0 SDK Example
'
' - VB6 SDK 연동환경 설정방법 안내 : http://blog.linkhub.co.kr/569
' - 업데이트 일자 : 2016-10-10
' - 연동 기술지원 연락처 : 1600-8536 / 070-4304-2991 (직통 / 정요한대리)
' - 연동 기술지원 이메일 : dev@linkhub.co.kr
'
' <테스트 연동개발 준비사항>
' 1) 30, 33번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
'    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
' 2) 팝빌 개발용 사이트(test.popbill.com)에 연동회원으로 가입합니다.
' 3) 전자세금계산서 발행을 위해 공인인증서를 등록합니다.
'    - 팝빌사이트 로그인 > [전자세금계산서] > [환경설정]
'      > [공인인증서 관리]
'    - 공인인증서 등록 팝업 URL (GetPopbillURL API)을 이용하여 등록
'
'=========================================================================

Option Explicit

'=========================================================================
' - 인증정보(링크아이디, 비밀키)는 파트너의 연동회원을 식별하는
'   인증에 사용되는 정보로 유출되지 않도록 주의하시기 바랍니다.
' - 상업용 전환이후에도 인증정보(링크아이디, 비밀키)는 변경되지 않습니다.
'=========================================================================

'링크아이디
Private Const LinkID = "TESTER"

'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'세금계산서 객체 생성
Private TaxinvoiceService As New PBTIService

'=========================================================================
' 팝빌 > 매입 문서함 팝업 URL을 반환합니다.
' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
'=========================================================================

Private Sub btn_GetURL_PBOX_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetURL(txtCorpNum.Text, txtUserID.Text, "PBOX")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 발행예정 세금계산서를 [승인]처리합니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    
    '메모
    memo = "발행예정 승인메모"
    
    Set Response = TaxinvoiceService.Accept(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 세금계산서에 첨부파일을 등록합니다.
' - [임시저장] 상태의 세금계산서만 파일을 첨부할수 있습니다.
' - 첨부파일은 최대 5개까지 등록할 수 있습니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.AttachFile(txtCorpNum.Text, KeyType, txtMgtKey.Text, FilePath, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    
End Sub

'=========================================================================
'1건의 전자명세서를 세금계산서에 첨부합니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '첨부할 전자명세서 종류코드, 121-거래명세서, 122-청구서, 123-견적서, 124-발주서, 125-입금표,126-영수증
    SubItemCode = 121
    
    '첨부할 전자명세서 관리번호
    SubMgtKey = "20151223-01"
        
    Set Response = TaxinvoiceService.AttachStatement(txtCorpNum.Text, KeyType, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
'[발행완료] 상태의 세금계산서를 [발행취소] 처리합니다.
' - [발행취소]는 국세청 전송전에만 가능합니다.
' - 발행취소된 세금계산서는 국세청에 전송되지 않습니다.
' - 발행취소 세금계산서에 기재된 문서관리번호를 재사용 하기 위해서는
'   삭제(Delete API)를 호출하여 [삭제] 처리 하셔야 합니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '메모
    memo = "발행취소 메모"
        
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    
End Sub

'=========================================================================
'[발행완료] 상태의 세금계산서를 [발행취소] 처리합니다.
' - [발행취소]는 국세청 전송전에만 가능합니다.
' - 발행취소된 세금계산서는 국세청에 전송되지 않습니다.
' - 발행취소 세금계산서에 기재된 문서관리번호를 재사용 하기 위해서는
'   삭제(Delete API)를 호출하여 [삭제] 처리 하셔야 합니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '메모
    memo = "발행 취소 메모"
    
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
'[발행완료] 상태의 세금계산서를 [발행취소] 처리합니다.
' - [발행취소]는 국세청 전송전에만 가능합니다.
' - 발행취소된 세금계산서는 국세청에 전송되지 않습니다.
' - 발행취소 세금계산서에 기재된 문서관리번호를 재사용 하기 위해서는
'   삭제(Delete API)를 호출하여 [삭제] 처리 하셔야 합니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '메모
    memo = "발행 취소 메모"
    
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 발행예정 세금계산서를 [취소] 처리 합니다.
' - [취소]된 세금계산서를 삭제(Delete API)하면 등록된 문서관리번호를
'   재사용할 수 있습니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '메모
    memo = "발행예정 취소 메모"
    
    Set Response = TaxinvoiceService.CancelSend(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 팝빌에 등록되어 있는 공인인증서의 만료일자를 확인합니다.
' - 공인인증서가 갱신/재발급/비밀번호 변경이 되는 경우 해당 인증서를
'   재등록 하셔야 정상적으로 API를 이용하실 수 있습니다.
'=========================================================================

Private Sub btnCertificateExpireDate_Click()
    Dim expireDate As String
        
    expireDate = TaxinvoiceService.GetCertificateExpireDate(txtCorpNum.Text)
    
    If expireDate = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "인증서만료일 : " + expireDate
 
End Sub

'=========================================================================
' 팝빌 회원아이디 중복여부를 확인합니다.
'=========================================================================

Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 해당 사업자의 파트너 연동회원 가입여부를 확인합니다.
' - LinkID는 인증정보로 설정되어 있는 링크아이디 값입니다.
'=========================================================================

Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse

    Set Response = TaxinvoiceService.CheckIsMember(txtCorpNum.Text, LinkID)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 세금계산서 관리번호 중복여부를 확인합니다.
' - 관리번호는 1~24자리로 숫자, 영문 '-', '_' 조합으로 구성할 수 있습니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.CheckMgtKeyInUse(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    
End Sub

'=========================================================================
' 1건의 전자세금계산서를 삭제합니다.
' - 세금계산서를 삭제해야만 문서관리번호(mgtKey)를 재사용할 수 있습니다.
' - 삭제가능한 문서 상태 : [임시저장], [발행취소], [발행예정 취소],
'   [발행예정 거부]
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.Delete(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 1건의 전자세금계산서를 삭제합니다.
' - 세금계산서를 삭제해야만 문서관리번호(mgtKey)를 재사용할 수 있습니다.
' - 삭제가능한 문서 상태 : [임시저장], [발행취소], [발행예정 취소],
'   [발행예정 거부]
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.Delete(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 1건의 전자세금계산서를 삭제합니다.
' - 세금계산서를 삭제해야만 문서관리번호(mgtKey)를 재사용할 수 있습니다.
' - 삭제가능한 문서 상태 : [임시저장], [발행취소], [발행예정 취소],
'   [발행예정 거부]
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.Delete(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 세금계산서에 첨부된 파일을 삭제합니다.
' - 파일을 식별하는 파일아이디는 첨부파일 목록(GetFileList API) 의 응답항목
'   중 파일아이디(AttachedFile) 값을 통해 확인할 수 있습니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.DeleteFile(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtFileID.Text, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 발행예정 세금계산서를 [거부]처리 합니다.
' - [거부]처리된 세금계산서를 삭제(Delete API)하면 등록된 문서관리번호를
'   재사용할 수 있습니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '메모
    memo = "발행예정 거부 메모"
    
    Set Response = TaxinvoiceService.Deny(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
'세금계산서에 첨부된 전자명세서 1건을 첨부해제합니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '첨부할 전자명세서 종류코드, 121-거래명세서, 122-청구서, 123-견적서, 124-발주서, 125-입금표,126-영수증
    SubItemCode = 121
    
    '첨부할 전자명세서 관리번호
    SubMgtKey = "20151223-01"
        
    Set Response = TaxinvoiceService.DetachStatement(txtCorpNum.Text, KeyType, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 잔여포인트를 확인합니다.
' - 과금방식이 파트너과금인 경우 파트너 잔여포인트(GetPartnerBalance API)
'   를 통해 확인하시기 바랍니다.
'=========================================================================
    
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = TaxinvoiceService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
End Sub

'=========================================================================
' 연동회원의 전자세금계산서 API 서비스 과금정보를 확인합니다.
'=========================================================================

Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    
    Set ChargeInfo = TaxinvoiceService.GetChargeInfo(txtCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "unitCost (발행단가) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (과금유형) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (과금제도) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원의 회사정보를 확인합니다.
'=========================================================================

Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    
    Set CorpInfo = TaxinvoiceService.GetCorpInfo(txtCorpNum.Text, txtUserID.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "ceoname(대표자성명) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName(상호) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr(주소) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType(업태) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass(종목) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
    
End Sub

'=========================================================================
' 1건의 세금계산서 상세항목을 확인합니다.
' - 응답항목에 대한 자세한 사항은 "[전자세금계산서 API 연동매뉴얼]
'   > 4.1 (세금)계산서 구성" 을 참조하시기 바랍니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set tiDetailInfo = TaxinvoiceService.GetDetailInfo(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
     
    If tiDetailInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
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

'=========================================================================
' 대용량 연계사업자 메일주소 목록을 반환합니다.
'=========================================================================

Private Sub btnGetEmailPublicKeys_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim email As Variant
    
    Set resultList = TaxinvoiceService.GetEmailPublicKeys(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If

    For Each email In resultList
        tmp = tmp + email + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 세금계산서 인쇄(공급받는자) URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    url = TaxinvoiceService.GetEPrintURL(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 세금계산서에 첨부된 파일의 목록을 확인합니다.
' - 응답항목 중 파일아이디(AttachedFile) 항목은 파일삭제(DeleteFile API)
'   호출시 이용할 수 있습니다.
'=========================================================================

Private Sub btnGetFiles_Click()
    Dim resultList As Collection
    Dim KeyType As MgtKeyType
    
    Set resultList = TaxinvoiceService.GetFiles(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
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
'1건의 세금계산서 상태/요약 정보를 확인합니다.
' - 세금계산서 상태정보(GetInfo API) 응답항목에 대한 자세한 정보는
'  "[전자세금계산서 API 연동매뉴얼] > 4.2. (세금)계산서 상태정보 구성"
'   을 참조하시기 바랍니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set tiInfo = TaxinvoiceService.GetInfo(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
     
    If tiInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
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
    tmp = tmp + "invoicerPrintYN : " + CStr(tiInfo.invoicerPrintYN) + vbCrLf
    
    tmp = tmp + "invoiceeCorpName : " + tiInfo.invoiceeCorpName + vbCrLf
    tmp = tmp + "invoiceeCorpNum : " + tiInfo.invoiceeCorpNum + vbCrLf
    tmp = tmp + "invoiceeMgtKey : " + tiInfo.invoiceeMgtKey + vbCrLf
    tmp = tmp + "invoiceePrintYN : " + CStr(tiInfo.invoiceePrintYN) + vbCrLf
    
    tmp = tmp + "trusteeCorpName : " + tiInfo.trusteeCorpName + vbCrLf
    tmp = tmp + "trusteeCorpNum : " + tiInfo.trusteeCorpNum + vbCrLf
    tmp = tmp + "trusteeMgtKey : " + tiInfo.trusteeMgtKey + vbCrLf
    tmp = tmp + "trusteePrintYN : " + CStr(tiInfo.trusteePrintYN) + vbCrLf
    
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

'=========================================================================
' 다량의 세금계산서 상태/요약 정보를 확인합니다. (최대 1000건)
' - 세금계산서 상태정보(GetInfos API) 응답항목에 대한 자세한 정보는
'  "[전자세금계산서 API 연동매뉴얼] > 4.2. (세금)계산서 상태정보 구성"
'  을 참조하시기 바랍니다.
'=========================================================================

Private Sub btnGetInfos_Click()
    Dim resultList As Collection
    Dim KeyType As MgtKeyType
    Dim KeyList As New Collection
    
    KeyType = SELL
    
    '세금계산서 문서관리번호 배열, 최대 1000건
    KeyList.Add "20161010-01"
    KeyList.Add "20161010-02"
    KeyList.Add "20161010-03"
    KeyList.Add "20161010-04"
    
    Set resultList = TaxinvoiceService.GetInfos(txtCorpNum.Text, KeyType, KeyList, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "ItemKey | StateCode | TaxType | WriteDate | RegDT | InvoicerPrintYN | InvoiceePrintYN | TrusteePrintYN " + vbCrLf
    
    Dim info As PBTIInfo
    
    For Each info In resultList
        tmp = tmp + info.itemKey + " | " + CStr(info.stateCode) + " | " + info.taxType + " | " + info.writeDate + " | " + info.regDT + " | "
        tmp = tmp + CStr(info.invoicerPrintYN) + " | " + CStr(info.invoiceePrintYN) + " | " + CStr(info.trusteePrintYN) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 세금계산서 상태 변경이력을 확인합니다.
' - 상태 변경이력 확인(GetLogs API) 응답항목에 대한 자세한 정보는
'   "[전자세금계산서 API 연동매뉴얼] > 3.6.4 상태 변경이력 확인"
'   을 참조하시기 바랍니다.
'=========================================================================

Private Sub btnGetLogs_Click()
    Dim resultList As Collection
    Dim KeyType As MgtKeyType
    
    Set resultList = TaxinvoiceService.GetLogs(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
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
' 공급받는자 메일링크 URL을 반환합니다.
' - 메일링크 URL은 유효시간이 존재하지 않습니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    url = TaxinvoiceService.GetMailURL(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 다수건의 전자세금계산서 인쇄팝업 URL을 반환합니다.
' 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
'=========================================================================

Private Sub btnGetMassPrintURL_Click()
    Dim url As String
    Dim KeyType As MgtKeyType
    Dim KeyList As New Collection
    
    '전자세금계산서 관리번호 배열
    KeyType = SELL
    KeyList.Add "20161010-01"
    KeyList.Add "20161010-02"
    KeyList.Add "20161010-03"
    KeyList.Add "20161010-04"
    KeyList.Add "20161010-05"
    
    url = TaxinvoiceService.GetMassPrintURL(txtCorpNum.Text, KeyType, KeyList, txtUserID.Text)
     
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    
End Sub
 
'=========================================================================
' 파트너의 잔여포인트를 확인합니다.
' - 과금방식이 연동과금인 경우 연동회원 잔여포인트(GetBalance API)를
'   이용하시기 바랍니다.
'=========================================================================

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = TaxinvoiceService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
    
End Sub

'=========================================================================
' 팝빌(www.popbill.com)에 로그인된 팝빌 URL을 반환합니다.
' - 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================

Private Sub btnGetPopbillURL_LOGIN_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
         MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 인감 및 첨부문서 등록 팝업 URL을 반환합니다.
' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
'=========================================================================

Private Sub btnGetPopbillURL_SEAL_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "SEAL")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
    
End Sub

'=========================================================================
' 1건의 전자세금계산서 보기 팝업 URL을 반환합니다.
' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    url = TaxinvoiceService.GetPopUpURL(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
    
End Sub

'=========================================================================
' 1건의 전자세금계산서 인쇄팝업 URL을 반환합니다.
' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    url = TaxinvoiceService.GetPrintURL(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 팝빌 > 매출 문서함 팝업 URL을 반환합니다.
' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
'=========================================================================

Private Sub btnGetURL_SBOX_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetURL(txtCorpNum.Text, txtUserID.Text, "SBOX")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 팝빌 > 임시(연동)문서함 팝업 URL을 반환합니다.
' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
'=========================================================================

Private Sub btnGetURL_TBOX_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetURL(txtCorpNum.Text, txtUserID.Text, "TBOX")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 팝빌 > 매출 문서작성 팝업 URL을 반환합니다.
' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
'=========================================================================

Private Sub btnGetURL_WRITE_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetURL(txtCorpNum.Text, txtUserID.Text, "WRITE")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' [임시저장] 상태의 세금계산서를 [발행]처리 합니다.
' - 발행(Issue API)를 호출하는 시점에서 포인트가 차감됩니다.
' - [발행완료] 세금계산서는 연동회원의 국세청 전송설정에 따라
'   익일/즉시전송 처리됩니다. 기본설정(익일전송)
' - 국세청 전송설정은 "팝빌 로그인" > [전자세금계산서] > [환경설정] >
'   [전자세금계산서 관리] > [국세청 전송 및 지연발행 설정] 탭에서
'   확인할 수 있습니다.
' - 국세청 전송정책에 대한 사항은 "[전자세금계산서 API 연동매뉴얼] >
'   1.4. 국세청 전송 정책" 을 참조하시기 바랍니다
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '메모
    memo = "메모"
    
    '공급받는자에게 전송되는 발행안내메일 제목, 미기재시 기본양식으로 전송
    emailSubject = ""
    
    '지연발행 강제여부, 기본값 - False
    '발행마감일이 지난 세금계산서를 발행하는 경우, 가산세가 부과될 수 있습니다.
    '지연발행 세금계산서를 신고해야 하는 경우 forceIssue 값을 True로 선언하여 발행(Issue API)을 호출할 수 있습니다.
    forceIssue = False
    
    Set Response = TaxinvoiceService.Issue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, emailSubject, forceIssue, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    
End Sub

'=========================================================================
' 역발행 요청받은 세금계산서를 [발행]처리 합니다.
' - 발행(Issue API)를 호출하는 시점에서 포인트가 차감됩니다.
' - [발행완료] 세금계산서는 연동회원의 국세청 전송설정에 따라
'    익일/즉시전송 처리됩니다. 기본설정(익일전송)
' - 국세청 전송설정은 "팝빌 로그인" > [전자세금계산서] > [환경설정] >
'   [전자세금계산서 관리] > [국세청 전송 및 지연발행 설정] 탭에서
'   확인할 수 있습니다.
' - 국세청 전송정책에 대한 사항은 "[전자세금계산서 API 연동매뉴얼] >
'   1.4. 국세청 전송 정책" 을 참조하시기 바랍니다
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '메모
    memo = "메모"
    
    '공급받는자에게 전송되는 발행안내메일 제목, 미기재시 기본양식으로 전송
    emailSubject = ""
    
    '지연발행 강제여부, 기본값 - False
    '발행마감일이 지난 세금계산서를 발행하는 경우, 가산세가 부과될 수 있습니다.
    '지연발행 세금계산서를 신고해야 하는 경우 forceIssue 값을 True로 선언하여 발행(Issue API)을 호출할 수 있습니다.
    forceIssue = False
    
    Set Response = TaxinvoiceService.Issue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, emailSubject, forceIssue, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 파트너의 연동회원으로 회원가입을 요청합니다.
'=========================================================================

Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '링크 아이디
    joinData.LinkID = LinkID
    
    '사업자번호, '-'제외, 10자리
    joinData.CorpNum = "1234567890"
    
    '대표자성명, 최대 30자
    joinData.ceoname = "대표자성명"
    
    '상호명, 최대 70자
    joinData.corpName = "회원상호"
    
    '주소, 최대 300자
    joinData.addr = "주소"
    
    '업태, 최대 40자
    joinData.bizType = "업태"
    
    '종목, 최대 40자
    joinData.bizClass = "종목"
    
    '아이디, 6자이상 20자 미만
    joinData.id = "userid"
    
    '비밀번호, 6자이상 20자 미만
    joinData.pwd = "pwd_must_be_long_enough"
    
    '담당자명, 최대 30자
    joinData.ContactName = "담당자성명"
    
    '담당자 연락처, 최대 20자
    joinData.ContactTEL = "02-999-9999"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.ContactHP = "010-1234-5678"
    
    '담당자 팩스번호, 최대 20자
    joinData.ContactFAX = "02-999-9998"
    
    '담당자 메일, 최대 70자
    joinData.ContactEmail = "test@test.com"
    
    Set Response = TaxinvoiceService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    
End Sub

'=========================================================================
' 연동회원의 담당자 목록을 확인합니다.
'=========================================================================

Private Sub btnListContact_Click()
    Dim resultList As Collection
        
    Set resultList = TaxinvoiceService.ListContact(txtCorpNum.Text, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
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

'=========================================================================
' 공인인증서 등록 URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================
    
Private Sub btnPopbillURL_CERT_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CERT")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
    
End Sub

'=========================================================================
' 연동회원 포인트 충전 URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================

Private Sub btnPopbillURL_CHRG_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 공급받는자에게 요청받은 역발행 세금계산서를 [거부]처리 합니다.
' - 세금계산서의 문서관리번호를 재사용하기 위해서는 삭제 (Delete API) 를
'   호출하여 [삭제] 처리해야 합니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '메모
    memo = "역발행 요청 거부 메모"
    
    Set Response = TaxinvoiceService.Refuse(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 담당자를 신규로 등록합니다.
'=========================================================================

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디, 6자 이상 20자 미만
    joinData.id = "testkorea"
    
    '비밀번호, 6자 이상 20자 미만
    joinData.pwd = "test@test.com"
    
    '담당자명, 최대 30자
    joinData.personName = "담당자명"
    
    '담당자 연락처
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호
    joinData.hp = "010-1234-1234"
    
    '담당자 메일주소
    joinData.email = "test@test.com"
    
    '담당자 팩스번호
    joinData.fax = "070-1234-1234"
    
    '회사조회 권한여부, true-회사조회 / false-개인조회
    joinData.searchAllAllowYN = True
    
    '관리자 권한여부
    joinData.mgrYN = False
    
    Set Response = TaxinvoiceService.RegistContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    
End Sub

'=========================================================================
' 1건의 세금계산서를 임시저장 합니다.
' - 세금계산서 임시저장(Register API) 호출후에는 발행(Issue API)을 호출해야만
'   국세청으로 전송됩니다.
' - 임시저장과 발행을 한번의 호출로 처리하는 즉시발행(RegistIssue API) 프로세스
'   연동을 권장합니다.
' - 세금계산서 항목별 정보는 "[전자세금계산서 API 연동매뉴얼] > 4.1. (세금)계산서
'   구성"을 참조하시기 바랍니다.
'=========================================================================
    
Private Sub btnRegister_Click()
    Dim Taxinvoice As New PBTaxinvoice
    Dim writeSpecification As Boolean
    
    '[필수] 작성일자, 표시형식 (yyyyMMdd) ex)20161010
    Taxinvoice.writeDate = "20161013"
    
    '[필수] 발행형태, [정발행, 역발행, 위수탁] 중 기재
    Taxinvoice.issueType = "정발행"
    
    '[필수] {정과금, 역과금} 중 기재, '역과금'은 역발행 프로세스에서만 이용가능
    '- 정과금(공급자 과금), 역과금(공급받는자 과금)
    Taxinvoice.chargeDirection = "정과금"
    
    '[필수] 영수/청구, [영수, 청구] 중 기재
    Taxinvoice.purposeType = "영수"
    
    '[필수] 발행시점, [직접발행, 승인시자동발행] 중 기재
    ' 발행예정(Send API) 프로세스를 구현하지 않는경우 '직접발행' 기재
    Taxinvoice.issueTiming = "직접발행"
    
    '[필수] 과세형태, [과세, 영세, 면세] 중 기재
    Taxinvoice.taxType = "과세"
    
    
    '=========================================================================
    '                              공급자 정보
    '=========================================================================
        
    '[필수] 공급자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoicerCorpNum = "1234567890"
    
    '[필수] 공급자 종사업자 식별번호. 필요시 숫자 4자리 기재
    Taxinvoice.invoicerTaxRegID = ""
    
    '[필수] 공급자 상호
    Taxinvoice.invoicerCorpName = "공급자 상호"
    
    '[필수] 공급자 문서관리번호, 1~24자리 (숫자, 영문, '-', '_') 조합으로
    '사업자 별로 중복되지 않도록 구성
    Taxinvoice.invoicerMgtKey = txtMgtKey.Text
    
    '[필수] 공급자 대표자 성명
    Taxinvoice.invoicerCEOName = "공급자 대표자 성명"
    
    '공급자 주소
    Taxinvoice.invoicerAddr = "공급자 주소"
    
    '공급자 업태
    Taxinvoice.invoicerBizType = "공급자 업태,업태2"
    
    '공급자 종목
    Taxinvoice.invoicerBizClass = "공급자 종목"
    
    '공급자 담당자명
    Taxinvoice.invoicerContactName = "공급자 담당자명"
    
    '공급자 담당자 메일주소
    Taxinvoice.invoicerEmail = "test@test.com"
    
    '공급자 담당자 연락처
    Taxinvoice.invoicerTEL = "070-7070-0707"
    
    '공급자 담당자 휴대폰번호
    Taxinvoice.invoicerHP = "010-000-2222"
    
    '정발행시 공급받는자에게 발행안내문자 전송여부
    '- 안내문자 전송기능 이용시 포인트가 차감됩니다.
    Taxinvoice.invoicerSMSSendYN = False
    
    
    '=========================================================================
    '                            공급받는자 정보
    '=========================================================================
        
    '[필수] 공급받는자 구분, [사업자, 개인, 외국인] 중 기재
    Taxinvoice.invoiceeType = "사업자"
    
    '[필수] 공급받는자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoiceeCorpNum = "8888888888"
    
    '[필수] 공급받는자 종사업자 식별번호. 필요시 숫자 4자리 기재
    Taxinvoice.invoiceeTaxRegID = ""
    
    '[필수] 공급자받는자 상호
    Taxinvoice.invoiceeCorpName = "공급받는자 상호"
    
    '[역발행시 필수] 공급받는자 문서관리번호(역발행시 필수)
    Taxinvoice.invoiceeMgtKey = ""
    
    '[필수] 공급받는자 대표자 성명
    Taxinvoice.invoiceeCEOName = "공급받는자 대표자 성명"
    
    '공급받는자 주소
    Taxinvoice.invoiceeAddr = "공급받는자 주소"
    
    '공급받는자 종목
    Taxinvoice.invoiceeBizClass = "공급받는자 업종"
    
    '공급받는자 업태
    Taxinvoice.invoiceeBizType = "공급받는자 업태"
    
    '공급받는자 담당자명
    Taxinvoice.invoiceeContactName1 = "공급받는자 담당자명"
    
    '공급받는자 담당자 메일주소
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    '공급받는자 담당자 연락처
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    
    '공급받는자 담당자 휴대폰번호
    Taxinvoice.invoiceeHP1 = "010-111-222"
    
    '역발행시 공급자에게 발행안내문자 전송여부
    Taxinvoice.invoiceeSMSSendYN = False
            
    
    '=========================================================================
    '                            세금계산서 정보
    '=========================================================================
    
    '[필수], 공급가액 합계
    Taxinvoice.supplyCostTotal = "200000"
    
    '[필수] 세액 합계
    Taxinvoice.taxTotal = "20000"
    
    '[필수] 합계금액, 공급가액 합계 + 세액합계
    Taxinvoice.totalAmount = "220000"
    
    '기재 상 '일련번호' 항목
    Taxinvoice.serialNum = "123"
    
    '기재 상 '권' 항목, 최대값 32767
    Taxinvoice.kwon = "1"
    
    '기재 상 '호' 항목, 최대값 32767
    Taxinvoice.ho = "1"
    
    '기재 상 '현금' 항목
    Taxinvoice.cash = ""
    
    '기재 상 '수표' 항목
    Taxinvoice.chkBill = ""
    
    '기재 상 '어음' 항목
    Taxinvoice.note = ""
    
    '기재 상 '외상미수금' 항목
    Taxinvoice.credit = ""
    
    '기재 상 '비고'항목
    Taxinvoice.remark1 = "비고1"
    Taxinvoice.remark2 = "비고2"
    Taxinvoice.remark3 = "비고3"
    
    '사업자등록증 이미지 첨부여부
    Taxinvoice.businessLicenseYN = False
    
    '통장사본 이미지 첨부여부
    Taxinvoice.bankBookYN = False
    

    '=========================================================================
    '         수정세금계산서 정보 (수정세금계산서 작성시에만 기재
    ' - 수정세금계산서 관련 정보는 연동매뉴얼 또는 개발가이드 링크 참조
    ' - [참고] 수정세금계산서 작성방법 안내 - http://blog.linkhub.co.kr/650
    '========================================================================='
    
    ' 수정사유코드, 수정사유에 따라 1~6중 선택기재
    Taxinvoice.modifyCode = ""
    
    ' 원본세금계산서의 ItemKey, 문서확인 (GetInfo API)의 응답결과(ItemKey 항목) 확인
    Taxinvoice.originalTaxinvoiceKey = ""
        
    
    '=========================================================================
    '                            상세항목(품목) 정보
    '=========================================================================
    
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20161010"   '거래일자  yyyyMMdd
    newDetail.itemName = "품명"         '품목명
    newDetail.spec = "규격"             '규격
    newDetail.qty = "1"                 '수량
    newDetail.unitCost = "100000"       '단가
    newDetail.supplyCost = "100000"     '공급가액
    newDetail.tax = "10000"             '세액
    newDetail.remark = "비고"           '비고
    
    Taxinvoice.detailList.Add newDetail
    
    Dim newDetail2 As New PBTIDetail
    newDetail2.serialNum = 2             '일련번호 1부터 순차 기재
    newDetail2.purchaseDT = "20161010"   '거래일자  yyyyMMdd
    newDetail2.itemName = "품명2"        '품목명
    newDetail2.spec = "규격"             '규격
    newDetail2.qty = "1"                 '수량
    newDetail2.unitCost = "100000"       '단가
    newDetail2.supplyCost = "100000"     '공급가액
    newDetail2.tax = "10000"             '세액
    newDetail2.remark = "비고"           '비고
    
    Taxinvoice.detailList.Add newDetail2
    
    
    '=========================================================================
    '                              추가담당자 정보
    ' - 세금계산서 발행안내 메일을 수신받을 공급받는자 담당자가 다수인 경우
    ' 담당자 정보를 추가하여 발행안내메일을 다수에게 전송할 수 있습니다.
    '=========================================================================
    
    
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                '일련번호, 1부터 순차기재
    newContact.ContactName = "담당자 성명"  '담당자명
    newContact.email = "test2@test.com"     '담당자 메일주소
    Taxinvoice.addContactList.Add newContact
    
    Dim newContact2 As New PBTIContact
    newContact2.serialNum = 2                '일련번호, 1부터 순차기재
    newContact2.ContactName = "담당자 성명"  '담당자명
    newContact2.email = "test2@test.com"     '담당자 메일주소
    
    Taxinvoice.addContactList.Add newContact2
    
    '거래명세서 동시작성 여부
    writeSpecification = False
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.Register(txtCorpNum.Text, Taxinvoice, writeSpecification, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    

End Sub

'=========================================================================
' 1건의 역발행 세금계산서를 [임시저장] 합니다.
' - 세금계산서 항목별 정보는 "[전자세금계산서 API 연동매뉴얼] > 4.1. (세금)계산서
'   구성"을 참조하시기 바랍니다.
'=========================================================================

Private Sub btnRegister_rev_Click()
    Dim Taxinvoice As New PBTaxinvoice
    
    '[필수] 작성일자, 표시형식 (yyyyMMdd) ex)20161010
    Taxinvoice.writeDate = "20161013"
    
    '[필수] 발행형태, [정발행, 역발행, 위수탁] 중 기재
    Taxinvoice.issueType = "역발행"
    
    '[필수] {정과금, 역과금} 중 기재, '역과금'은 역발행 프로세스에서만 이용가능
    '- 정과금(공급자 과금), 역과금(공급받는자 과금)
    Taxinvoice.chargeDirection = "정과금"
    
    '[필수] 영수/청구, [영수, 청구] 중 기재
    Taxinvoice.purposeType = "영수"
    
    '[필수] 발행시점, [직접발행, 승인시자동발행] 중 기재
    ' 발행예정(Send API) 프로세스를 구현하지 않는경우 '직접발행' 기재
    Taxinvoice.issueTiming = "직접발행"
    
    '[필수] 과세형태, [과세, 영세, 면세] 중 기재
    Taxinvoice.taxType = "과세"
    
    
    '=========================================================================
    '                              공급자 정보
    '=========================================================================
        
    '[필수] 공급자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoicerCorpNum = "8888888888"
    
    '[필수] 공급자 종사업자 식별번호. 필요시 숫자 4자리 기재
    Taxinvoice.invoicerTaxRegID = ""
    
    '[필수] 공급자 상호
    Taxinvoice.invoicerCorpName = "공급자 상호"
    
    '[필수] 공급자 문서관리번호, 1~24자리 (숫자, 영문, '-', '_') 조합으로
    '사업자 별로 중복되지 않도록 구성
    Taxinvoice.invoicerMgtKey = ""
    
    '[필수] 공급자 대표자 성명
    Taxinvoice.invoicerCEOName = "공급자 대표자 성명"
    
    '공급자 주소
    Taxinvoice.invoicerAddr = "공급자 주소"
    
    '공급자 업태
    Taxinvoice.invoicerBizType = "공급자 업태,업태2"
    
    '공급자 종목
    Taxinvoice.invoicerBizClass = "공급자 종목"
    
    '공급자 담당자명
    Taxinvoice.invoicerContactName = "공급자 담당자명"
    
    '공급자 담당자 메일주소
    Taxinvoice.invoicerEmail = "test@test.com"
    
    '공급자 담당자 연락처
    Taxinvoice.invoicerTEL = "070-7070-0707"
    
    '공급자 담당자 휴대폰번호
    Taxinvoice.invoicerHP = "010-000-2222"
    
    '정발행시 공급받는자에게 발행안내문자 전송여부
    '- 안내문자 전송기능 이용시 포인트가 차감됩니다.
    Taxinvoice.invoicerSMSSendYN = False
    
    
    '=========================================================================
    '                            공급받는자 정보
    '=========================================================================
        
    '[필수] 공급받는자 구분, [사업자, 개인, 외국인] 중 기재
    Taxinvoice.invoiceeType = "사업자"
    
    '[필수] 공급받는자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoiceeCorpNum = "1234567890"
    
    '[필수] 공급받는자 종사업자 식별번호. 필요시 숫자 4자리 기재
    Taxinvoice.invoiceeTaxRegID = ""
    
    '[필수] 공급자받는자 상호
    Taxinvoice.invoiceeCorpName = "공급받는자 상호"
    
    '[역발행시 필수] 공급받는자 문서관리번호(역발행시 필수)
    Taxinvoice.invoiceeMgtKey = txtMgtKey.Text
    
    '[필수] 공급받는자 대표자 성명
    Taxinvoice.invoiceeCEOName = "공급받는자 대표자 성명"
    
    '공급받는자 주소
    Taxinvoice.invoiceeAddr = "공급받는자 주소"
    
    '공급받는자 종목
    Taxinvoice.invoiceeBizClass = "공급받는자 업종"
    
    '공급받는자 업태
    Taxinvoice.invoiceeBizType = "공급받는자 업태"
    
    '공급받는자 담당자명
    Taxinvoice.invoiceeContactName1 = "공급받는자 담당자명"
    
    '공급받는자 담당자 메일주소
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    '공급받는자 담당자 연락처
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    
    '공급받는자 담당자 휴대폰번호
    Taxinvoice.invoiceeHP1 = "010-111-222"
    
    '역발행시 공급자에게 발행안내문자 전송여부
    Taxinvoice.invoiceeSMSSendYN = False
            
    
    '=========================================================================
    '                            세금계산서 정보
    '=========================================================================
    
    '[필수], 공급가액 합계
    Taxinvoice.supplyCostTotal = "200000"
    
    '[필수] 세액 합계
    Taxinvoice.taxTotal = "20000"
    
    '[필수] 합계금액, 공급가액 합계 + 세액합계
    Taxinvoice.totalAmount = "220000"
    
    '기재 상 '일련번호' 항목
    Taxinvoice.serialNum = "123"
    
    '기재 상 '권' 항목, 최대값 32767
    Taxinvoice.kwon = "1"
    
    '기재 상 '호' 항목, 최대값 32767
    Taxinvoice.ho = "1"
    
    '기재 상 '현금' 항목
    Taxinvoice.cash = ""
    
    '기재 상 '수표' 항목
    Taxinvoice.chkBill = ""
    
    '기재 상 '어음' 항목
    Taxinvoice.note = ""
    
    '기재 상 '외상미수금' 항목
    Taxinvoice.credit = ""
    
    '기재 상 '비고'항목
    Taxinvoice.remark1 = "비고1"
    Taxinvoice.remark2 = "비고2"
    Taxinvoice.remark3 = "비고3"
    
    '사업자등록증 이미지 첨부여부
    Taxinvoice.businessLicenseYN = False
    
    '통장사본 이미지 첨부여부
    Taxinvoice.bankBookYN = False
    

    '=========================================================================
    '         수정세금계산서 정보 (수정세금계산서 작성시에만 기재
    ' - 수정세금계산서 관련 정보는 연동매뉴얼 또는 개발가이드 링크 참조
    ' - [참고] 수정세금계산서 작성방법 안내 - http://blog.linkhub.co.kr/650
    '========================================================================='
    
    ' 수정사유코드, 수정사유에 따라 1~6중 선택기재
    Taxinvoice.modifyCode = ""
    
    ' 원본세금계산서의 ItemKey, 문서확인 (GetInfo API)의 응답결과(ItemKey 항목) 확인
    Taxinvoice.originalTaxinvoiceKey = ""
        
    
    '=========================================================================
    '                            상세항목(품목) 정보
    '=========================================================================
    
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20161010"   '거래일자  yyyyMMdd
    newDetail.itemName = "품명"         '품목명
    newDetail.spec = "규격"             '규격
    newDetail.qty = "1"                 '수량
    newDetail.unitCost = "100000"       '단가
    newDetail.supplyCost = "100000"     '공급가액
    newDetail.tax = "10000"             '세액
    newDetail.remark = "비고"           '비고
    
    Taxinvoice.detailList.Add newDetail
    
    Dim newDetail2 As New PBTIDetail
    newDetail2.serialNum = 2             '일련번호 1부터 순차 기재
    newDetail2.purchaseDT = "20161010"   '거래일자  yyyyMMdd
    newDetail2.itemName = "품명2"        '품목명
    newDetail2.spec = "규격"             '규격
    newDetail2.qty = "1"                 '수량
    newDetail2.unitCost = "100000"       '단가
    newDetail2.supplyCost = "100000"     '공급가액
    newDetail2.tax = "10000"             '세액
    newDetail2.remark = "비고"           '비고
    
    Taxinvoice.detailList.Add newDetail2
    
    
    '=========================================================================
    '                              추가담당자 정보
    ' - 세금계산서 발행안내 메일을 수신받을 공급받는자 담당자가 다수인 경우
    ' 담당자 정보를 추가하여 발행안내메일을 다수에게 전송할 수 있습니다.
    '=========================================================================
    
    
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                '일련번호, 1부터 순차기재
    newContact.ContactName = "담당자 성명"  '담당자명
    newContact.email = "test2@test.com"     '담당자 메일주소
    Taxinvoice.addContactList.Add newContact
    
    Dim newContact2 As New PBTIContact
    newContact2.serialNum = 2                '일련번호, 1부터 순차기재
    newContact2.ContactName = "담당자 성명"  '담당자명
    newContact2.email = "test2@test.com"     '담당자 메일주소
    
    Taxinvoice.addContactList.Add newContact2
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.Register(txtCorpNum.Text, Taxinvoice, False, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 1건의 세금계산서를 즉시발행 처리합니다.
' - 세금계산서 항목별 정보는 "[전자세금계산서 API 연동매뉴얼] > 4.1. (세금)계산서
'   구성"을 참조하시기 바랍니다.
'=========================================================================

Private Sub btnRegistIssue_Click()
    Dim Taxinvoice As New PBTaxinvoice
        
   '[필수] 작성일자, 표시형식 (yyyyMMdd) ex)20161010
    Taxinvoice.writeDate = "20161013"
    
    '[필수] 발행형태, [정발행, 역발행, 위수탁] 중 기재
    Taxinvoice.issueType = "정발행"
    
    '[필수] {정과금, 역과금} 중 기재, '역과금'은 역발행 프로세스에서만 이용가능
    '- 정과금(공급자 과금), 역과금(공급받는자 과금)
    Taxinvoice.chargeDirection = "정과금"
    
    '[필수] 영수/청구, [영수, 청구] 중 기재
    Taxinvoice.purposeType = "영수"
    
    '[필수] 발행시점, [직접발행, 승인시자동발행] 중 기재
    ' 발행예정(Send API) 프로세스를 구현하지 않는경우 '직접발행' 기재
    Taxinvoice.issueTiming = "직접발행"
    
    '[필수] 과세형태, [과세, 영세, 면세] 중 기재
    Taxinvoice.taxType = "과세"
    
    
    '=========================================================================
    '                              공급자 정보
    '=========================================================================
        
    '[필수] 공급자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoicerCorpNum = "1234567890"
    
    '[필수] 공급자 종사업자 식별번호. 필요시 숫자 4자리 기재
    Taxinvoice.invoicerTaxRegID = ""
    
    '[필수] 공급자 상호
    Taxinvoice.invoicerCorpName = "공급자 상호"
    
    '[필수] 공급자 문서관리번호, 1~24자리 (숫자, 영문, '-', '_') 조합으로
    '사업자 별로 중복되지 않도록 구성
    Taxinvoice.invoicerMgtKey = txtMgtKey.Text
    
    '[필수] 공급자 대표자 성명
    Taxinvoice.invoicerCEOName = "공급자 대표자 성명"
    
    '공급자 주소
    Taxinvoice.invoicerAddr = "공급자 주소"
    
    '공급자 업태
    Taxinvoice.invoicerBizType = "공급자 업태,업태2"
    
    '공급자 종목
    Taxinvoice.invoicerBizClass = "공급자 종목"
    
    '공급자 담당자명
    Taxinvoice.invoicerContactName = "공급자 담당자명"
    
    '공급자 담당자 메일주소
    Taxinvoice.invoicerEmail = "test@test.com"
    
    '공급자 담당자 연락처
    Taxinvoice.invoicerTEL = "070-7070-0707"
    
    '공급자 담당자 휴대폰번호
    Taxinvoice.invoicerHP = "010-000-2222"
    
    '정발행시 공급받는자에게 발행안내문자 전송여부
    '- 안내문자 전송기능 이용시 포인트가 차감됩니다.
    Taxinvoice.invoicerSMSSendYN = False
    
    
    '=========================================================================
    '                            공급받는자 정보
    '=========================================================================
        
    '[필수] 공급받는자 구분, [사업자, 개인, 외국인] 중 기재
    Taxinvoice.invoiceeType = "사업자"
    
    '[필수] 공급받는자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoiceeCorpNum = "8888888888"
    
    '[필수] 공급받는자 종사업자 식별번호. 필요시 숫자 4자리 기재
    Taxinvoice.invoiceeTaxRegID = ""
    
    '[필수] 공급자받는자 상호
    Taxinvoice.invoiceeCorpName = "공급받는자 상호"
    
    '[역발행시 필수] 공급받는자 문서관리번호(역발행시 필수)
    Taxinvoice.invoiceeMgtKey = ""
    
    '[필수] 공급받는자 대표자 성명
    Taxinvoice.invoiceeCEOName = "공급받는자 대표자 성명"
    
    '공급받는자 주소
    Taxinvoice.invoiceeAddr = "공급받는자 주소"
    
    '공급받는자 종목
    Taxinvoice.invoiceeBizClass = "공급받는자 업종"
    
    '공급받는자 업태
    Taxinvoice.invoiceeBizType = "공급받는자 업태"
    
    '공급받는자 담당자명
    Taxinvoice.invoiceeContactName1 = "공급받는자 담당자명"
    
    '공급받는자 담당자 메일주소
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    '공급받는자 담당자 연락처
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    
    '공급받는자 담당자 휴대폰번호
    Taxinvoice.invoiceeHP1 = "010-111-222"
    
    '역발행시 공급자에게 발행안내문자 전송여부
    Taxinvoice.invoiceeSMSSendYN = False
            
    
    '=========================================================================
    '                            세금계산서 정보
    '=========================================================================
    
    '[필수], 공급가액 합계
    Taxinvoice.supplyCostTotal = "200000"
    
    '[필수] 세액 합계
    Taxinvoice.taxTotal = "20000"
    
    '[필수] 합계금액, 공급가액 합계 + 세액합계
    Taxinvoice.totalAmount = "220000"
    
    '기재 상 '일련번호' 항목
    Taxinvoice.serialNum = "123"
    
    '기재 상 '권' 항목, 최대값 32767
    Taxinvoice.kwon = "1"
    
    '기재 상 '호' 항목, 최대값 32767
    Taxinvoice.ho = "1"
    
    '기재 상 '현금' 항목
    Taxinvoice.cash = ""
    
    '기재 상 '수표' 항목
    Taxinvoice.chkBill = ""
    
    '기재 상 '어음' 항목
    Taxinvoice.note = ""
    
    '기재 상 '외상미수금' 항목
    Taxinvoice.credit = ""
    
    '기재 상 '비고'항목
    Taxinvoice.remark1 = "비고1"
    Taxinvoice.remark2 = "비고2"
    Taxinvoice.remark3 = "비고3"
    
    '사업자등록증 이미지 첨부여부
    Taxinvoice.businessLicenseYN = False
    
    '통장사본 이미지 첨부여부
    Taxinvoice.bankBookYN = False
    

    '=========================================================================
    '         수정세금계산서 정보 (수정세금계산서 작성시에만 기재
    ' - 수정세금계산서 관련 정보는 연동매뉴얼 또는 개발가이드 링크 참조
    ' - [참고] 수정세금계산서 작성방법 안내 - http://blog.linkhub.co.kr/650
    '========================================================================='
    
    ' 수정사유코드, 수정사유에 따라 1~6중 선택기재
    Taxinvoice.modifyCode = ""
    
    ' 원본세금계산서의 ItemKey, 문서확인 (GetInfo API)의 응답결과(ItemKey 항목) 확인
    Taxinvoice.originalTaxinvoiceKey = ""
        
    
    '=========================================================================
    '                            상세항목(품목) 정보
    '=========================================================================
    
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20161010"   '거래일자  yyyyMMdd
    newDetail.itemName = "품명"         '품목명
    newDetail.spec = "규격"             '규격
    newDetail.qty = "1"                 '수량
    newDetail.unitCost = "100000"       '단가
    newDetail.supplyCost = "100000"     '공급가액
    newDetail.tax = "10000"             '세액
    newDetail.remark = "비고"           '비고
    
    Taxinvoice.detailList.Add newDetail
    
    Dim newDetail2 As New PBTIDetail
    newDetail2.serialNum = 2             '일련번호 1부터 순차 기재
    newDetail2.purchaseDT = "20161010"   '거래일자  yyyyMMdd
    newDetail2.itemName = "품명2"        '품목명
    newDetail2.spec = "규격"             '규격
    newDetail2.qty = "1"                 '수량
    newDetail2.unitCost = "100000"       '단가
    newDetail2.supplyCost = "100000"     '공급가액
    newDetail2.tax = "10000"             '세액
    newDetail2.remark = "비고"           '비고
    
    Taxinvoice.detailList.Add newDetail2
    
    
    '=========================================================================
    '                              추가담당자 정보
    ' - 세금계산서 발행안내 메일을 수신받을 공급받는자 담당자가 다수인 경우
    ' 담당자 정보를 추가하여 발행안내메일을 다수에게 전송할 수 있습니다.
    '=========================================================================
    
    
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                '일련번호, 1부터 순차기재
    newContact.ContactName = "담당자 성명"  '담당자명
    newContact.email = "test2@test.com"     '담당자 메일주소
    Taxinvoice.addContactList.Add newContact
    
    Dim newContact2 As New PBTIContact
    newContact2.serialNum = 2                '일련번호, 1부터 순차기재
    newContact2.ContactName = "담당자 성명"  '담당자명
    newContact2.email = "test2@test.com"     '담당자 메일주소
    
    Taxinvoice.addContactList.Add newContact2
        
        
    '거래명세서 동시작성 여부
    Taxinvoice.writeSpecification = False
    
    '지연발행 강제여부(forceIssue)
    '발행마감일이 지난 세금계산서를 발행하는 경우, 가산세가 부과될 수 있습니다.
    '가산세가 부과되더라도 발행을 해야하는 경우에는 forceIssue의 값을
    'true로 선언하여 발행(Issue API)를 호출하시면 됩니다.
    Taxinvoice.forceIssue = False
    
    '메모
    Taxinvoice.memo = ""
    
    '발행안내 메일제목, 공백처리시 기본제목으로 전송
    Taxinvoice.emailSubject = ""
    
    '거래명세서 동시작성시 거래명세서 관리번호, 미기재시 세금계산서 관리번호로 자동작성
    Taxinvoice.dealInvoiceMgtKey = ""
    
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.RegistIssue(txtCorpNum.Text, Taxinvoice, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    
End Sub

'=========================================================================
' 공급받는자가 공급자에게 1건의 역발행 세금계산서를 요청합니다.
' - 역발행 세금계산서 프로세스를 구현하기 위해서는 공급자/공급받는자가 모두
'   팝빌에 회원이여야 합니다.
' - 역발행 요청후 공급자가 [발행] 처리시 포인트가 차감되며 역발행
'   세금계산서 항목중 과금방향(ChargeDirection) 에 기재한 값에 따라
'   정과금(공급자과금) 또는 역과금(공급받는자과금) 처리됩니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '메모
    memo = "역발행 요청 메모"
    
    Set Response = TaxinvoiceService.Request(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 역발행 세금계산서를 [취소] 처리합니다.
' - [취소]한 세금계산서의 문서관리번호를 재사용하기 위해서는 삭제 (Delete API)
'   를 호출해야 합니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '메모
    memo = "역발행 요청 취소 메모"
    
    Set Response = TaxinvoiceService.CancelRequest(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 검색조건을 사용하여 세금계산서 목록을 조회합니다.
' - 응답항목에 대한 자세한 사항은 "[전자세금계산서 API 연동매뉴얼] >
'   4.2. (세금)계산서 상태정보 구성" 을 참조하시기 바랍니다.
'=========================================================================

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
    Dim TaxRegIDType As String
    Dim TaxRegID As String
    Dim TaxRegIDYN As String
    Dim QString As String
    
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
    
    '[필수] 일자유형, R-등록일시 W-작성일자 I-발행일시 중 택1
    DType = "W"
    
    '[필수] 시작일자, yyyyMMdd
    SDate = "20160901"
    
    '[필수] 종료일자, yyyyMMdd
    EDate = "20161031"
    
    '전송상태값 배열, 미기재시 전체상태조회, 문서상태값 3자리숫자 작성
    '2,3번째 와일드카드 가능
    State.Add "3**"
    State.Add "6**"
    
    '문서유형 배열, N-일반 M-수정 중 선택, 미기재시 전체조회
    TType.Add "N"
    TType.Add "M"
    
    '과세형태 배열, T-과세, N-면세 Z-영세 중 선택, 미기재시 전체조회
    taxType.Add "T"
    taxType.Add "N"
    taxType.Add "Z"
    
    '지연발행 여부, 0-정상발행분만 조회 1-지연발행분만조회, 공백처리시 전체조회
    LateOnly = ""
    
    '페이지 번호
    Page = 1
    
    '페이지 목록개수, 최대 1000건
    PerPage = 10
    
    '정렬방향, D-내림차순(기본값), A-오름차순
    Order = "D"
    
    '종사업장번호 유형 S-공급자, B-공급받는자, T-수탁자
    TaxRegIDType = "S"
    
    '종사업장번호, 콤마(,)로 구분하여 구성 ex) 0001,0002
    TaxRegID = ""
    
    '종사업장 유무, 공백-전체조회, 0-종사업장번호 없는경우만 조회, 1-종사업장번호 조건 조회
    TaxRegIDYN = ""
    
    '거래처 조회, 거래처 상호 또는 거래처 사업자등록번호 조회, 공백처리시 전체조회
    QString = ""
    
    Set tiSearchList = TaxinvoiceService.Search(txtCorpNum.Text, KeyType, DType, SDate, EDate, State, TType, _
                        taxType, LateOnly, Page, PerPage, Order, TaxRegIDType, TaxRegID, TaxRegIDYN, QString)
     
    If tiSearchList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    tmp = "code : " + CStr(tiSearchList.code) + vbCrLf
    tmp = tmp + "total : " + CStr(tiSearchList.total) + vbCrLf
    tmp = tmp + "perPage : " + CStr(tiSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum : " + CStr(tiSearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount : " + CStr(tiSearchList.pageCount) + vbCrLf
    tmp = tmp + "message : " + tiSearchList.message + vbCrLf + vbCrLf
    
    tmp = tmp + "itemKey | stateCode | TaxTye | writeDate | regDT | lateIssueYN | invoicerCorpNum | invoicerCorpName | invoicerPrintYn | invoiceeCorpNum | invoiceeCorpName | invoiceePrintYN | " + _
                "issueType | supplyCostTotal | taxTotal | trusteePrintYN " + vbCrLf
            
    Dim info As PBTIInfo
    
    For Each info In tiSearchList.list
        tmp = tmp + info.itemKey + " | "
        tmp = tmp + CStr(info.stateCode) + " | "
        tmp = tmp + info.taxType + " | "
        tmp = tmp + info.writeDate + " | "
        tmp = tmp + info.regDT + " | "
        tmp = tmp + CStr(info.lateIssueYN) + " | "
        tmp = tmp + info.invoicerCorpNum + " | "
        tmp = tmp + info.invoicerCorpName + " | "
        tmp = tmp + CStr(info.invoicerPrintYN) + " | "
        tmp = tmp + info.invoiceeCorpNum + " | "
        tmp = tmp + info.invoiceeCorpName + " | "
        tmp = tmp + CStr(info.invoiceePrintYN) + " | "
        tmp = tmp + info.issueType + " | "
        tmp = tmp + info.supplyCostTotal + " | "
        tmp = tmp + info.taxTotal + vbCrLf
    Next
    
    MsgBox tmp
       
End Sub

'=========================================================================
' 1건의 [임시저장] 상태의 세금계산서를 [발행예정] 처리합니다.
' - 발행예정이란 공급자와 공급받는자 사이에 세금계산서 확인 후 발행하는
'   방법입니다.
' - "[전자세금계산서 API 연동매뉴얼] > 1.3.1. 정발행 프로세스 흐름도
'   > 다. 임시저장 발행예정" 의 프로세스를 참조하시기 바랍니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '발행예정 안내메일 제목, 공백처리시 기본제목으로 전송
    emailSubject = ""
    
    '메모
    memo = "발행예정 메모"
    
    Set Response = TaxinvoiceService.Send(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, emailSubject, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 발행 안내메일을 재전송합니다.
' - 메일내용중 전자세금계산서 [보기] 버튼이 동작하지 않는 경우,
'   키보드 왼쪽 Shift 키를 누르고 버튼을 클릭해보시기 바랍니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '수신자 메일주소
    receiveEmail = "test@test.com"
    
    Set Response = TaxinvoiceService.SendEmail(txtCorpNum.Text, KeyType, txtMgtKey.Text, receiveEmail, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 전자세금계산서를 팩스전송합니다.
' - 팩스 전송 요청시 포인트가 차감됩니다. (전송실패시 환불처리)
' - 전송내역 확인은 "팝빌 로그인" > [문자 팩스] > [팩스] > [전송내역]
'   메뉴에서 전송결과를 확인할 수 있습니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
        
    '발신번호
    senderNum = "07043042991"
    
    '수신번호
    receiveNum = "111-222-4444"
        
    Set Response = TaxinvoiceService.SendFax(txtCorpNum.Text, KeyType, txtMgtKey.Text, senderNum, receiveNum, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 알림문자를 전송합니다. (단문/SMS- 한글 최대 45자)
' - 알림문자 전송시 포인트가 차감됩니다. (전송실패시 환불처리)
' - 전송내역 확인은 "팝빌 로그인" > [문자 팩스] > [전송내역] 탭에서
'   전송결과를 확인할 수 있습니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    ' 발신번호, [참고] 발신번호 세칙규정 - http://blog.linkhub.co.kr/3064
    senderNum = "07043042991"
    
    ' 수신번호
    receiveNum = "010-111-222"
    
    ' 메시지 내용, 최대 90Byte(한글 45자), 길이를 초과한 내용은 삭제되어 전송됩니다.
    Contents = "링크허브에서 세금계산서를 발행하였습니다. 메일확인 바랍니다."
    
    Set Response = TaxinvoiceService.SendSMS(txtCorpNum.Text, KeyType, txtMgtKey.Text, _
                                senderNum, receiveNum, Contents, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub


'=========================================================================
' [발행완료] 상태의 세금계산서를 국세청으로 즉시전송합니다.
' - 국세청 즉시전송을 호출하지 않은 세금계산서는 발행일 기준 익일 오후 3시에
'   팝빌 시스템에서 일괄적으로 국세청으로 전송합니다.
' - 익일전송시 전송일이 법정공휴일인 경우 다음 영업일에 전송됩니다.
' - 국세청 전송에 관한 사항은 "[전자세금계산서 API 연동매뉴얼] > 1.4 국세청
'   전송 정책" 을 참조하시기 바랍니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.SendToNTS(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 전자세금계산서 발행단가를 확인합니다.
'=========================================================================
    
Private Sub btnUnitCost_Click()
    Dim unitCost As Double
    
    unitCost = TaxinvoiceService.GetUnitCost(txtCorpNum.Text)
    
    If unitCost < 0 Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "발행단가 : " + CStr(unitCost)
End Sub

'=========================================================================
' [임시저장] 상태의 세금계산서의 항목을 수정합니다.
' - 세금계산서 항목별 정보는 "[전자세금계산서 API 연동매뉴얼] > 4.1. (세금)계산서
'   구성"을 참조하시기 바랍니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Dim Taxinvoice As New PBTaxinvoice
    
    '[필수] 작성일자, 표시형식 (yyyyMMdd) ex)20161010
    Taxinvoice.writeDate = "20161010"
    
    '[필수] 발행형태, [정발행, 역발행, 위수탁] 중 기재
    Taxinvoice.issueType = "정발행"
    
    '[필수] {정과금, 역과금} 중 기재, '역과금'은 역발행 프로세스에서만 이용가능
    '- 정과금(공급자 과금), 역과금(공급받는자 과금)
    Taxinvoice.chargeDirection = "정과금"
    
    '[필수] 영수/청구, [영수, 청구] 중 기재
    Taxinvoice.purposeType = "영수"
    
    '[필수] 발행시점, [직접발행, 승인시자동발행] 중 기재
    ' 발행예정(Send API) 프로세스를 구현하지 않는경우 '직접발행' 기재
    Taxinvoice.issueTiming = "직접발행"
    
    '[필수] 과세형태, [과세, 영세, 면세] 중 기재
    Taxinvoice.taxType = "과세"
    
    
    '=========================================================================
    '                              공급자 정보
    '=========================================================================
        
    '[필수] 공급자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoicerCorpNum = "1234567890"
    
    '[필수] 공급자 종사업자 식별번호. 필요시 숫자 4자리 기재
    Taxinvoice.invoicerTaxRegID = ""
    
    '[필수] 공급자 상호
    Taxinvoice.invoicerCorpName = "공급자 상호_수정"
    
    '[필수] 공급자 문서관리번호, 1~24자리 (숫자, 영문, '-', '_') 조합으로
    '사업자 별로 중복되지 않도록 구성
    Taxinvoice.invoicerMgtKey = txtMgtKey.Text
    
    '[필수] 공급자 대표자 성명
    Taxinvoice.invoicerCEOName = "공급자 대표자 성명_수정"
    
    '공급자 주소
    Taxinvoice.invoicerAddr = "공급자 주소"
    
    '공급자 업태
    Taxinvoice.invoicerBizType = "공급자 업태,업태2"
    
    '공급자 종목
    Taxinvoice.invoicerBizClass = "공급자 종목"
    
    '공급자 담당자명
    Taxinvoice.invoicerContactName = "공급자 담당자명"
    
    '공급자 담당자 메일주소
    Taxinvoice.invoicerEmail = "test@test.com"
    
    '공급자 담당자 연락처
    Taxinvoice.invoicerTEL = "070-7070-0707"
    
    '공급자 담당자 휴대폰번호
    Taxinvoice.invoicerHP = "010-000-2222"
    
    '정발행시 공급받는자에게 발행안내문자 전송여부
    '- 안내문자 전송기능 이용시 포인트가 차감됩니다.
    Taxinvoice.invoicerSMSSendYN = False
    
    
    '=========================================================================
    '                            공급받는자 정보
    '=========================================================================
        
    '[필수] 공급받는자 구분, [사업자, 개인, 외국인] 중 기재
    Taxinvoice.invoiceeType = "사업자"
    
    '[필수] 공급받는자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoiceeCorpNum = "8888888888"
    
    '[필수] 공급받는자 종사업자 식별번호. 필요시 숫자 4자리 기재
    Taxinvoice.invoiceeTaxRegID = ""
    
    '[필수] 공급자받는자 상호
    Taxinvoice.invoiceeCorpName = "공급받는자 상호"
    
    '[역발행시 필수] 공급받는자 문서관리번호(역발행시 필수)
    Taxinvoice.invoiceeMgtKey = ""
    
    '[필수] 공급받는자 대표자 성명
    Taxinvoice.invoiceeCEOName = "공급받는자 대표자 성명"
    
    '공급받는자 주소
    Taxinvoice.invoiceeAddr = "공급받는자 주소"
    
    '공급받는자 종목
    Taxinvoice.invoiceeBizClass = "공급받는자 업종"
    
    '공급받는자 업태
    Taxinvoice.invoiceeBizType = "공급받는자 업태"
    
    '공급받는자 담당자명
    Taxinvoice.invoiceeContactName1 = "공급받는자 담당자명"
    
    '공급받는자 담당자 메일주소
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    '공급받는자 담당자 연락처
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    
    '공급받는자 담당자 휴대폰번호
    Taxinvoice.invoiceeHP1 = "010-111-222"
    
    '역발행시 공급자에게 발행안내문자 전송여부
    Taxinvoice.invoiceeSMSSendYN = False
            
    
    '=========================================================================
    '                            세금계산서 정보
    '=========================================================================
    
    '[필수], 공급가액 합계
    Taxinvoice.supplyCostTotal = "200000"
    
    '[필수] 세액 합계
    Taxinvoice.taxTotal = "20000"
    
    '[필수] 합계금액, 공급가액 합계 + 세액합계
    Taxinvoice.totalAmount = "220000"
    
    '기재 상 '일련번호' 항목
    Taxinvoice.serialNum = "123"
    
    '기재 상 '권' 항목, 최대값 32767
    Taxinvoice.kwon = "1"
    
    '기재 상 '호' 항목, 최대값 32767
    Taxinvoice.ho = "1"
    
    '기재 상 '현금' 항목
    Taxinvoice.cash = ""
    
    '기재 상 '수표' 항목
    Taxinvoice.chkBill = ""
    
    '기재 상 '어음' 항목
    Taxinvoice.note = ""
    
    '기재 상 '외상미수금' 항목
    Taxinvoice.credit = ""
    
    '기재 상 '비고'항목
    Taxinvoice.remark1 = "비고1"
    Taxinvoice.remark2 = "비고2"
    Taxinvoice.remark3 = "비고3"
    
    '사업자등록증 이미지 첨부여부
    Taxinvoice.businessLicenseYN = False
    
    '통장사본 이미지 첨부여부
    Taxinvoice.bankBookYN = False
    

    '=========================================================================
    '         수정세금계산서 정보 (수정세금계산서 작성시에만 기재
    ' - 수정세금계산서 관련 정보는 연동매뉴얼 또는 개발가이드 링크 참조
    ' - [참고] 수정세금계산서 작성방법 안내 - http://blog.linkhub.co.kr/650
    '========================================================================='
    
    ' 수정사유코드, 수정사유에 따라 1~6중 선택기재
    Taxinvoice.modifyCode = ""
    
    ' 원본세금계산서의 ItemKey, 문서확인 (GetInfo API)의 응답결과(ItemKey 항목) 확인
    Taxinvoice.originalTaxinvoiceKey = ""
        
    
    '=========================================================================
    '                            상세항목(품목) 정보
    '=========================================================================
    
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20161010"   '거래일자  yyyyMMdd
    newDetail.itemName = "품명"         '품목명
    newDetail.spec = "규격"             '규격
    newDetail.qty = "1"                 '수량
    newDetail.unitCost = "100000"       '단가
    newDetail.supplyCost = "100000"     '공급가액
    newDetail.tax = "10000"             '세액
    newDetail.remark = "비고"           '비고
    
    Taxinvoice.detailList.Add newDetail
    
    Dim newDetail2 As New PBTIDetail
    newDetail2.serialNum = 2             '일련번호 1부터 순차 기재
    newDetail2.purchaseDT = "20161010"   '거래일자  yyyyMMdd
    newDetail2.itemName = "품명2"        '품목명
    newDetail2.spec = "규격"             '규격
    newDetail2.qty = "1"                 '수량
    newDetail2.unitCost = "100000"       '단가
    newDetail2.supplyCost = "100000"     '공급가액
    newDetail2.tax = "10000"             '세액
    newDetail2.remark = "비고"           '비고
    
    Taxinvoice.detailList.Add newDetail2
    
    
    '=========================================================================
    '                              추가담당자 정보
    ' - 세금계산서 발행안내 메일을 수신받을 공급받는자 담당자가 다수인 경우
    ' 담당자 정보를 추가하여 발행안내메일을 다수에게 전송할 수 있습니다.
    '=========================================================================
    
    
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                '일련번호, 1부터 순차기재
    newContact.ContactName = "담당자 성명"  '담당자명
    newContact.email = "test2@test.com"     '담당자 메일주소
    Taxinvoice.addContactList.Add newContact
    
    Dim newContact2 As New PBTIContact
    newContact2.serialNum = 2                '일련번호, 1부터 순차기재
    newContact2.ContactName = "담당자 성명"  '담당자명
    newContact2.email = "test2@test.com"     '담당자 메일주소
    
    Taxinvoice.addContactList.Add newContact2
    
    '거래명세서 동시작성여부
    writeSpecification = False
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.Update(txtCorpNum.Text, KeyType, txtMgtKey.Text, Taxinvoice, writeSpecification, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' [임시저장] 상태의 세금계산서의 항목을 수정합니다.
' - 세금계산서 항목별 정보는 "[전자세금계산서 API 연동매뉴얼] > 4.1. (세금)계산서
'   구성"을 참조하시기 바랍니다.
'=========================================================================

Private Sub btnUpdate_rev_Click()
    Dim KeyType As MgtKeyType
    
    KeyType = BUY
    
    Dim Taxinvoice As New PBTaxinvoice
    
    '[필수] 작성일자, 표시형식 (yyyyMMdd) ex)20161010
    Taxinvoice.writeDate = "20161010"
    
    '[필수] 발행형태, [정발행, 역발행, 위수탁] 중 기재
    Taxinvoice.issueType = "역발행"
    
    '[필수] {정과금, 역과금} 중 기재, '역과금'은 역발행 프로세스에서만 이용가능
    '- 정과금(공급자 과금), 역과금(공급받는자 과금)
    Taxinvoice.chargeDirection = "정과금"
    
    '[필수] 영수/청구, [영수, 청구] 중 기재
    Taxinvoice.purposeType = "영수"
    
    '[필수] 발행시점, [직접발행, 승인시자동발행] 중 기재
    ' 발행예정(Send API) 프로세스를 구현하지 않는경우 '직접발행' 기재
    Taxinvoice.issueTiming = "직접발행"
    
    '[필수] 과세형태, [과세, 영세, 면세] 중 기재
    Taxinvoice.taxType = "과세"
    
    
    '=========================================================================
    '                              공급자 정보
    '=========================================================================
        
    '[필수] 공급자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoicerCorpNum = "8888888888"
    
    '[필수] 공급자 종사업자 식별번호. 필요시 숫자 4자리 기재
    Taxinvoice.invoicerTaxRegID = ""
    
    '[필수] 공급자 상호
    Taxinvoice.invoicerCorpName = "공급자 상호_수정"
    
    '[필수] 공급자 문서관리번호, 1~24자리 (숫자, 영문, '-', '_') 조합으로
    '사업자 별로 중복되지 않도록 구성
    Taxinvoice.invoicerMgtKey = ""
    
    '[필수] 공급자 대표자 성명
    Taxinvoice.invoicerCEOName = "공급자 대표자 성명_수정"
    
    '공급자 주소
    Taxinvoice.invoicerAddr = "공급자 주소"
    
    '공급자 업태
    Taxinvoice.invoicerBizType = "공급자 업태,업태2"
    
    '공급자 종목
    Taxinvoice.invoicerBizClass = "공급자 종목"
    
    '공급자 담당자명
    Taxinvoice.invoicerContactName = "공급자 담당자명"
    
    '공급자 담당자 메일주소
    Taxinvoice.invoicerEmail = "test@test.com"
    
    '공급자 담당자 연락처
    Taxinvoice.invoicerTEL = "070-7070-0707"
    
    '공급자 담당자 휴대폰번호
    Taxinvoice.invoicerHP = "010-000-2222"
    
    '정발행시 공급받는자에게 발행안내문자 전송여부
    '- 안내문자 전송기능 이용시 포인트가 차감됩니다.
    Taxinvoice.invoicerSMSSendYN = False
    
    
    '=========================================================================
    '                            공급받는자 정보
    '=========================================================================
        
    '[필수] 공급받는자 구분, [사업자, 개인, 외국인] 중 기재
    Taxinvoice.invoiceeType = "사업자"
    
    '[필수] 공급받는자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoiceeCorpNum = "1234567890"
    
    '[필수] 공급받는자 종사업자 식별번호. 필요시 숫자 4자리 기재
    Taxinvoice.invoiceeTaxRegID = ""
    
    '[필수] 공급자받는자 상호
    Taxinvoice.invoiceeCorpName = "공급받는자 상호"
    
    '[역발행시 필수] 공급받는자 문서관리번호(역발행시 필수)
    Taxinvoice.invoiceeMgtKey = txtMgtKey.Text
    
    '[필수] 공급받는자 대표자 성명
    Taxinvoice.invoiceeCEOName = "공급받는자 대표자 성명"
    
    '공급받는자 주소
    Taxinvoice.invoiceeAddr = "공급받는자 주소"
    
    '공급받는자 종목
    Taxinvoice.invoiceeBizClass = "공급받는자 업종"
    
    '공급받는자 업태
    Taxinvoice.invoiceeBizType = "공급받는자 업태"
    
    '공급받는자 담당자명
    Taxinvoice.invoiceeContactName1 = "공급받는자 담당자명"
    
    '공급받는자 담당자 메일주소
    Taxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    '공급받는자 담당자 연락처
    Taxinvoice.invoiceeTEL1 = "070-1234-1234"
    
    '공급받는자 담당자 휴대폰번호
    Taxinvoice.invoiceeHP1 = "010-111-222"
    
    '역발행시 공급자에게 발행안내문자 전송여부
    Taxinvoice.invoiceeSMSSendYN = False
            
    
    '=========================================================================
    '                            세금계산서 정보
    '=========================================================================
    
    '[필수], 공급가액 합계
    Taxinvoice.supplyCostTotal = "200000"
    
    '[필수] 세액 합계
    Taxinvoice.taxTotal = "20000"
    
    '[필수] 합계금액, 공급가액 합계 + 세액합계
    Taxinvoice.totalAmount = "220000"
    
    '기재 상 '일련번호' 항목
    Taxinvoice.serialNum = "123"
    
    '기재 상 '권' 항목, 최대값 32767
    Taxinvoice.kwon = "1"
    
    '기재 상 '호' 항목, 최대값 32767
    Taxinvoice.ho = "1"
    
    '기재 상 '현금' 항목
    Taxinvoice.cash = ""
    
    '기재 상 '수표' 항목
    Taxinvoice.chkBill = ""
    
    '기재 상 '어음' 항목
    Taxinvoice.note = ""
    
    '기재 상 '외상미수금' 항목
    Taxinvoice.credit = ""
    
    '기재 상 '비고'항목
    Taxinvoice.remark1 = "비고1"
    Taxinvoice.remark2 = "비고2"
    Taxinvoice.remark3 = "비고3"
    
    '사업자등록증 이미지 첨부여부
    Taxinvoice.businessLicenseYN = False
    
    '통장사본 이미지 첨부여부
    Taxinvoice.bankBookYN = False         '통장사본 이미지 첨부시 설정.
    

    '=========================================================================
    '         수정세금계산서 정보 (수정세금계산서 작성시에만 기재
    ' - 수정세금계산서 관련 정보는 연동매뉴얼 또는 개발가이드 링크 참조
    ' - [참고] 수정세금계산서 작성방법 안내 - http://blog.linkhub.co.kr/650
    '========================================================================='
    
    ' 수정사유코드, 수정사유에 따라 1~6중 선택기재
    Taxinvoice.modifyCode = ""
    
    ' 원본세금계산서의 ItemKey, 문서확인 (GetInfo API)의 응답결과(ItemKey 항목) 확인
    Taxinvoice.originalTaxinvoiceKey = ""
        
    
    '=========================================================================
    '                            상세항목(품목) 정보
    '=========================================================================
    
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20161010"   '거래일자  yyyyMMdd
    newDetail.itemName = "품명"         '품목명
    newDetail.spec = "규격"             '규격
    newDetail.qty = "1"                 '수량
    newDetail.unitCost = "100000"       '단가
    newDetail.supplyCost = "100000"     '공급가액
    newDetail.tax = "10000"             '세액
    newDetail.remark = "비고"           '비고
    
    Taxinvoice.detailList.Add newDetail
    
    Dim newDetail2 As New PBTIDetail
    newDetail2.serialNum = 2             '일련번호 1부터 순차 기재
    newDetail2.purchaseDT = "20161010"   '거래일자  yyyyMMdd
    newDetail2.itemName = "품명2"        '품목명
    newDetail2.spec = "규격"             '규격
    newDetail2.qty = "1"                 '수량
    newDetail2.unitCost = "100000"       '단가
    newDetail2.supplyCost = "100000"     '공급가액
    newDetail2.tax = "10000"             '세액
    newDetail2.remark = "비고"           '비고
    
    Taxinvoice.detailList.Add newDetail2
    
    
    '=========================================================================
    '                              추가담당자 정보
    ' - 세금계산서 발행안내 메일을 수신받을 공급받는자 담당자가 다수인 경우
    ' 담당자 정보를 추가하여 발행안내메일을 다수에게 전송할 수 있습니다.
    '=========================================================================
    
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                '일련번호, 1부터 순차기재
    newContact.ContactName = "담당자 성명"  '담당자명
    newContact.email = "test2@test.com"     '담당자 메일주소
    Taxinvoice.addContactList.Add newContact
    
    Dim newContact2 As New PBTIContact
    newContact2.serialNum = 2                '일련번호, 1부터 순차기재
    newContact2.ContactName = "담당자 성명"  '담당자명
    newContact2.email = "test2@test.com"     '담당자 메일주소
    
    Taxinvoice.addContactList.Add newContact2
    
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.Update(txtCorpNum.Text, KeyType, txtMgtKey.Text, Taxinvoice, False, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 담당자 정보를 수정합니다.
'=========================================================================

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자명
    joinData.personName = "담당자명_수정"
    
    '연락처
    joinData.tel = "070-1234-1234"
    
    '휴대폰번호
    joinData.hp = "010-1234-1234"
    
    '이메일 주소
    joinData.email = "test@test.com"
    
    '팩스번호
    joinData.fax = "070-1234-1234"
    
    '전체조회여부, Ture-회사조회, False-개인조
    joinData.searchAllAllowYN = True
    
    '관리자 권한여부
    joinData.mgrYN = False
                
    Set Response = TaxinvoiceService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 회사정보를 수정합니다
'=========================================================================

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '대표자명, 최대 30자
    CorpInfo.ceoname = "대표자"
    
    '상호, 최대 70자
    CorpInfo.corpName = "상호"
    
    ' 주소, 최대 300자
    CorpInfo.addr = "서울특별시"
    
    '업태, 최대 40자
    CorpInfo.bizType = "업태"
    
    '종목, 최대 40자
    CorpInfo.bizClass = "종목"
    
    Set Response = TaxinvoiceService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub


Private Sub Form_Load()
    ' 모듈 초기화
    TaxinvoiceService.Initialize LinkID, SecretKey
    
    ' 연동환경 설정값 True(개발용), False(상업용), 상업용 전환시 False로 변경.
    TaxinvoiceService.IsTest = True
        
    cboMgtKeyType.AddItem "SELL"
    cboMgtKeyType.AddItem "BUY"
    cboMgtKeyType.AddItem "TRUSTEE"
End Sub

