VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "팝빌 세금계산서 SDK 예제"
   ClientHeight    =   12705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19065
   LinkTopic       =   "Form1"
   ScaleHeight     =   12705
   ScaleWidth      =   19065
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton btnUpdateemailconfig 
      Caption         =   "알림메일 전송설정 수정"
      Height          =   390
      Index           =   3
      Left            =   7440
      TabIndex        =   86
      Top             =   11040
      Width           =   2085
   End
   Begin VB.CommandButton btnListemailconfig 
      Caption         =   "알림메일 전송목록 조회"
      Height          =   390
      Index           =   2
      Left            =   7440
      TabIndex        =   85
      Top             =   10560
      Width           =   2085
   End
   Begin VB.CommandButton btnAssignmgtkey 
      Caption         =   "관리번호 할당"
      Height          =   390
      Index           =   1
      Left            =   5400
      TabIndex        =   84
      Top             =   11040
      Width           =   1965
   End
   Begin VB.Frame Frame17 
      Caption         =   "파트너과금 포인트"
      Height          =   1935
      Index           =   1
      Left            =   6720
      TabIndex        =   79
      Top             =   840
      Width           =   2415
      Begin VB.CommandButton btnGetPartnerURL_CHRG 
         Caption         =   "포인트 충전 URL"
         Height          =   410
         Left            =   120
         TabIndex        =   83
         Top             =   840
         Width           =   2150
      End
      Begin VB.CommandButton btnGetPartnerBalance 
         Caption         =   "파트너 잔여포인트 확인"
         Height          =   410
         Left            =   120
         TabIndex        =   82
         Top             =   360
         Width           =   2150
      End
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "목록 조회"
      Height          =   390
      Left            =   2950
      TabIndex        =   72
      Top             =   11040
      Width           =   1845
   End
   Begin VB.CommandButton btnUpdateCorpInfo 
      Caption         =   "회사정보 수정"
      Height          =   410
      Left            =   9360
      TabIndex        =   68
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton btnUpdateContact 
      Caption         =   "담당자 정보 수정"
      Height          =   410
      Left            =   13800
      TabIndex        =   66
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton btnListContact 
      Caption         =   "담당자 목록 조회"
      Height          =   410
      Left            =   13800
      TabIndex        =   65
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Frame Frame15 
      Caption         =   "회사정보 관련"
      Height          =   1935
      Left            =   9240
      TabIndex        =   63
      Top             =   840
      Width           =   2055
      Begin VB.CommandButton btnGetCorpInfo 
         Caption         =   "회사정보 조회"
         Height          =   410
         Left            =   120
         TabIndex        =   67
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton btnCheckID 
      Caption         =   "ID 중복 확인"
      Height          =   410
      Left            =   480
      TabIndex        =   62
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   " 팝빌 기본 API "
      Height          =   2295
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   18495
      Begin VB.Frame Frame17 
         Caption         =   "연동과금 포인트"
         Height          =   1935
         Index           =   0
         Left            =   4440
         TabIndex        =   78
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   81
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   " 포인트 충전 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   80
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " 공인인증서 관련"
         Height          =   1935
         Left            =   11160
         TabIndex        =   12
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton btnCheckCertValidation 
            Caption         =   "인증서 유효성 확인"
            Height          =   375
            Left            =   120
            TabIndex        =   89
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton btnGetTaxCertURL 
            Caption         =   " 인증서 등록 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   76
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton btnCertificateExpireDate 
            Caption         =   "인증서 만료일 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보"
         Height          =   1935
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1695
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "가입 여부 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "회원 가입"
            Height          =   410
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " 포인트 관련"
         Height          =   1935
         Left            =   1920
         TabIndex        =   8
         Top             =   240
         Width           =   2415
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "과금정보 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   75
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "요금 단가 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "담당자 관련"
         Height          =   1935
         Left            =   13440
         TabIndex        =   7
         Top             =   240
         Width           =   2055
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "담당자 추가"
            Height          =   410
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " 팝빌 기본 URL"
         Height          =   1935
         Left            =   15600
         TabIndex        =   5
         Top             =   240
         Width           =   2655
         Begin VB.CommandButton btnGetSealURL 
            Caption         =   "인감 및 첨부문서 등록 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   77
            Top             =   840
            Width           =   2415
         End
         Begin VB.CommandButton btnGetAccessURL 
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
   Begin VB.Frame Frame18 
      Caption         =   " (권장) 즉시요청 프로세스"
      Height          =   3255
      Left            =   10080
      TabIndex        =   90
      Top             =   4920
      Width           =   3615
      Begin VB.CommandButton btnRegistRequest 
         BackColor       =   &H00FFFFC0&
         Caption         =   "즉시요청"
         Height          =   420
         Left            =   1560
         Style           =   1  '그래픽
         TabIndex        =   96
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton btnIssue_rev_sub 
         BackColor       =   &H00C0C0FF&
         Caption         =   "발행"
         Height          =   420
         Left            =   330
         Style           =   1  '그래픽
         TabIndex        =   95
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton btnDelete_rev_sub 
         Caption         =   "삭제"
         Height          =   420
         Left            =   2520
         Style           =   1  '그래픽
         TabIndex        =   94
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton btnCancelIssue_rev_sub 
         BackColor       =   &H00C0C0FF&
         Caption         =   "발행취소"
         Height          =   420
         Left            =   330
         Style           =   1  '그래픽
         TabIndex        =   93
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton btnRequestCancel_sub 
         BackColor       =   &H00FFFFC0&
         Caption         =   "요청취소"
         Height          =   420
         Left            =   2520
         Style           =   1  '그래픽
         TabIndex        =   92
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton btnRefuse_sub 
         BackColor       =   &H00C0C0FF&
         Caption         =   "거부"
         Height          =   420
         Left            =   1440
         Style           =   1  '그래픽
         TabIndex        =   91
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
         BackStyle       =   0  '투명
         Caption         =   "즉시요청"
         Height          =   180
         Left            =   480
         TabIndex        =   98
         Top             =   600
         Width           =   720
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  '투명하지 않음
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
      Caption         =   " 세금계산서 관련 기능"
      Height          =   9225
      Left            =   240
      TabIndex        =   13
      Top             =   3120
      Width           =   18495
      Begin VB.Frame Frame16 
         Caption         =   " (권장) 즉시발행 프로세스"
         Height          =   3255
         Left            =   720
         TabIndex        =   69
         Top             =   1800
         Width           =   3255
         Begin VB.CommandButton btnRegistIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "즉시발행"
            Height          =   495
            Left            =   360
            Style           =   1  '그래픽
            TabIndex        =   107
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton btnCancelIssue_sub 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행취소"
            Height          =   495
            Left            =   360
            Style           =   1  '그래픽
            TabIndex        =   71
            Top             =   2110
            Width           =   975
         End
         Begin VB.CommandButton btnDelete_sub 
            Caption         =   "삭제"
            Height          =   495
            Left            =   1920
            Style           =   1  '그래픽
            TabIndex        =   70
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
            BackStyle       =   1  '투명하지 않음
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
         Caption         =   "유통사업자메일 목록"
         Height          =   375
         Left            =   10080
         TabIndex        =   60
         Top             =   240
         Width           =   1965
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   240
         Top             =   5280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame14 
         Caption         =   " 보기/인쇄"
         Height          =   2760
         Left            =   9600
         TabIndex        =   55
         Top             =   6120
         Width           =   3210
         Begin VB.CommandButton btnGetEPrintUrl 
            Caption         =   "공급받는자 인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   61
            Top             =   1260
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPopUpURL 
            Caption         =   "세금계산서 보기 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   59
            Top             =   390
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPrintURL 
            Caption         =   "공급자 인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   58
            Top             =   825
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMassPrintURL 
            Caption         =   "대량 인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   57
            Top             =   1710
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMailURL 
            Caption         =   "세금계산서 메일링크 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   56
            Top             =   2160
            Width           =   2745
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   " 기타 URL "
         Height          =   2295
         Left            =   12960
         TabIndex        =   50
         Top             =   6120
         Width           =   2265
         Begin VB.CommandButton btnGetURL_TBOX 
            Caption         =   "임시 문서함"
            Height          =   390
            Left            =   210
            TabIndex        =   54
            Top             =   390
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_SBOX 
            Caption         =   "매출 문서함"
            Height          =   390
            Left            =   210
            TabIndex        =   53
            Top             =   825
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_PBOX 
            Caption         =   "매입 문서함"
            Height          =   390
            Left            =   195
            TabIndex        =   52
            Top             =   1260
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_WRITE 
            Caption         =   "매출 문서작성"
            Height          =   390
            Left            =   195
            TabIndex        =   51
            Top             =   1710
            Width           =   1845
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   " 부가 서비스"
         Height          =   2415
         Left            =   4920
         TabIndex        =   48
         Top             =   6120
         Width           =   4545
         Begin VB.CommandButton btnSendFAX 
            Caption         =   "팩스 전송"
            Height          =   375
            Left            =   240
            TabIndex        =   88
            Top             =   1320
            Width           =   1965
         End
         Begin VB.CommandButton btnSendSMS 
            Caption         =   "문자 전송"
            Height          =   375
            Left            =   240
            TabIndex        =   87
            Top             =   840
            Width           =   1965
         End
         Begin VB.CommandButton btnDetachStatement 
            Caption         =   "전자명세서 첨부해제"
            Height          =   390
            Left            =   2280
            TabIndex        =   74
            Top             =   840
            Width           =   2085
         End
         Begin VB.CommandButton btnAttachStatement 
            Caption         =   "전자명세서 첨부"
            Height          =   390
            Left            =   2280
            TabIndex        =   73
            Top             =   390
            Width           =   2085
         End
         Begin VB.CommandButton btnSendEmail 
            Caption         =   "이메일 전송"
            Height          =   390
            Left            =   240
            TabIndex        =   49
            Top             =   390
            Width           =   1965
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   " 정보 확인"
         Height          =   2775
         Left            =   2520
         TabIndex        =   43
         Top             =   6120
         Width           =   2265
         Begin VB.CommandButton btnGetDetailInfo 
            Caption         =   "상세정보 확인"
            Height          =   390
            Left            =   195
            TabIndex        =   47
            Top             =   1320
            Width           =   1845
         End
         Begin VB.CommandButton btnGetLogs 
            Caption         =   "상태 변경이력"
            Height          =   390
            Left            =   195
            TabIndex        =   46
            Top             =   2280
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfos 
            Caption         =   "상태 대량 확인"
            Height          =   390
            Left            =   210
            TabIndex        =   45
            Top             =   825
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfo 
            Caption         =   "상태 확인"
            Height          =   390
            Left            =   210
            TabIndex        =   44
            Top             =   390
            Width           =   1845
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   " 첨부파일 "
         Height          =   2280
         Left            =   120
         TabIndex        =   38
         Top             =   6135
         Width           =   2265
         Begin VB.CommandButton btnDeleteFile 
            Caption         =   "파일 삭제"
            Height          =   390
            Left            =   210
            TabIndex        =   42
            Top             =   1650
            Width           =   1845
         End
         Begin VB.TextBox txtFileID 
            Height          =   330
            Left            =   210
            TabIndex        =   41
            Text            =   "파일아이디"
            Top             =   1245
            Width           =   1845
         End
         Begin VB.CommandButton btnGetFiles 
            Caption         =   "첨부 목록"
            Height          =   390
            Left            =   210
            TabIndex        =   40
            Top             =   795
            Width           =   1845
         End
         Begin VB.CommandButton btnAttachFile 
            Caption         =   "파일 첨부"
            Height          =   390
            Left            =   210
            TabIndex        =   39
            Top             =   345
            Width           =   1845
         End
      End
      Begin VB.CommandButton btnSendToNTS 
         BackColor       =   &H00C0C0FF&
         Caption         =   "국세청 즉시 전송"
         Height          =   375
         Left            =   2400
         Style           =   1  '그래픽
         TabIndex        =   37
         Top             =   5160
         Width           =   4200
      End
      Begin VB.Frame Frame9 
         Caption         =   " 임시저장 역발행 프로세스 "
         Height          =   3255
         Left            =   13680
         TabIndex        =   21
         Top             =   1800
         Width           =   4095
         Begin VB.CommandButton btnRefuse 
            BackColor       =   &H00C0C0FF&
            Caption         =   "거부"
            Height          =   420
            Left            =   1320
            Style           =   1  '그래픽
            TabIndex        =   36
            Top             =   1800
            Width           =   855
         End
         Begin VB.CommandButton btnRequestCancel 
            BackColor       =   &H00FFFFC0&
            Caption         =   "요청취소"
            Height          =   420
            Left            =   2760
            Style           =   1  '그래픽
            TabIndex        =   35
            Top             =   1200
            Width           =   855
         End
         Begin VB.CommandButton btnCancelIssue_rev 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행취소"
            Height          =   420
            Left            =   360
            Style           =   1  '그래픽
            TabIndex        =   34
            Top             =   2520
            Width           =   855
         End
         Begin VB.CommandButton btnDelete_rev 
            Caption         =   "삭제"
            Height          =   420
            Left            =   2760
            Style           =   1  '그래픽
            TabIndex        =   33
            Top             =   2520
            Width           =   855
         End
         Begin VB.CommandButton btnIssue_rev 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행"
            Height          =   420
            Left            =   360
            Style           =   1  '그래픽
            TabIndex        =   32
            Top             =   1800
            Width           =   855
         End
         Begin VB.CommandButton btnRequest 
            BackColor       =   &H00FFFFC0&
            Caption         =   "역)발행요청"
            Height          =   420
            Left            =   320
            Style           =   1  '그래픽
            TabIndex        =   31
            Top             =   1200
            Width           =   1920
         End
         Begin VB.CommandButton btnUpdate_rev 
            BackColor       =   &H00FFFFC0&
            Caption         =   "수정"
            Height          =   375
            Left            =   2475
            Style           =   1  '그래픽
            TabIndex        =   29
            Top             =   465
            Width           =   855
         End
         Begin VB.CommandButton btnRegister_rev 
            BackColor       =   &H00FFFFC0&
            Caption         =   "등록"
            Height          =   375
            Left            =   1515
            Style           =   1  '그래픽
            TabIndex        =   28
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
            BackStyle       =   0  '투명
            Caption         =   "임시저장"
            Height          =   180
            Left            =   675
            TabIndex        =   30
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
         Caption         =   " 임시저장 발행 프로세스"
         Height          =   3255
         Left            =   4200
         TabIndex        =   20
         Top             =   1800
         Width           =   4695
         Begin VB.CommandButton btnRegister 
            BackColor       =   &H00C0C0FF&
            Caption         =   "등록"
            Height          =   375
            Left            =   1305
            Style           =   1  '그래픽
            TabIndex        =   26
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton btnUpdate 
            BackColor       =   &H00C0C0FF&
            Caption         =   "수정"
            Height          =   375
            Left            =   2265
            Style           =   1  '그래픽
            TabIndex        =   25
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton btnDelete 
            Caption         =   "삭제"
            Height          =   375
            Left            =   3345
            Style           =   1  '그래픽
            TabIndex        =   24
            Top             =   2280
            Width           =   975
         End
         Begin VB.CommandButton btnCancelIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행취소"
            Height          =   375
            Left            =   360
            Style           =   1  '그래픽
            TabIndex        =   23
            Top             =   2280
            Width           =   975
         End
         Begin VB.CommandButton btnIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행"
            Height          =   495
            Left            =   360
            Style           =   1  '그래픽
            TabIndex        =   22
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "임시저장"
            Height          =   180
            Left            =   465
            TabIndex        =   27
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
         Caption         =   "관리번호 사용여부 확인"
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
         Left            =   5160
         TabIndex        =   101
         Top             =   960
         Width           =   3615
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   ": 공급받는자 처리"
            Height          =   180
            Left            =   2040
            TabIndex        =   103
            Top             =   270
            Width           =   1440
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   ": 공급자 처리"
            Height          =   180
            Left            =   480
            TabIndex        =   102
            Top             =   270
            Width           =   1080
         End
         Begin VB.Shape Shape8 
            BackColor       =   &H00C0C0FF&
            BorderColor     =   &H00404040&
            FillColor       =   &H00C0C0FF&
            FillStyle       =   0  '단색
            Height          =   255
            Left            =   120
            Top             =   240
            Width           =   255
         End
         Begin VB.Shape Shape7 
            BackColor       =   &H00FFFFC0&
            BorderColor     =   &H00404040&
            FillColor       =   &H00FFFFC0&
            FillStyle       =   0  '단색
            Height          =   255
            Left            =   1680
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   " 역발행 프로세스"
         Height          =   4935
         Left            =   9600
         TabIndex        =   99
         Top             =   840
         Width           =   8415
         Begin VB.Frame Frame22 
            Height          =   615
            Left            =   4680
            TabIndex        =   104
            Top             =   120
            Width           =   3615
            Begin VB.Shape Shape10 
               BackColor       =   &H00FFFFC0&
               BorderColor     =   &H00404040&
               FillColor       =   &H00FFFFC0&
               FillStyle       =   0  '단색
               Height          =   255
               Left            =   1680
               Top             =   240
               Width           =   255
            End
            Begin VB.Shape Shape9 
               BackColor       =   &H00C0C0FF&
               BorderColor     =   &H00404040&
               FillColor       =   &H00C0C0FF&
               FillStyle       =   0  '단색
               Height          =   255
               Left            =   120
               Top             =   240
               Width           =   255
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   ": 공급자 처리"
               Height          =   180
               Left            =   480
               TabIndex        =   106
               Top             =   270
               Width           =   1080
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   ": 공급받는자 처리"
               Height          =   180
               Left            =   2040
               TabIndex        =   105
               Top             =   270
               Width           =   1440
            End
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   " 정발행 프로세스"
         Height          =   4935
         Left            =   480
         TabIndex        =   97
         Top             =   840
         Width           =   8775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "문서관리번호( MgtKey) : "
         Height          =   180
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   2115
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "임시저장"
      Height          =   180
      Left            =   8040
      TabIndex        =   100
      Top             =   5400
      Width           =   720
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  '투명하지 않음
      FillColor       =   &H00E0E0E0&
      Height          =   660
      Left            =   10200
      Top             =   4440
      Width           =   3360
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
' - 업데이트 일자 : 2019-02-12
' - 연동 기술지원 연락처 : 1600-9854 / 070-4304-2991
' - 연동 기술지원 이메일 : code@linkhub.co.kr
'
' <테스트 연동개발 준비사항>
' 1) 30, 33번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
'    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
' 2) 팝빌 개발용 사이트(test.popbill.com)에 연동회원으로 가입합니다.
' 3) 전자세금계산서 발행을 위해 공인인증서를 등록합니다.
'    - 팝빌사이트 로그인 > [전자세금계산서] > [환경설정]
'      > [공인인증서 관리]
'    - 공인인증서 등록 팝업 URL (GetTaxCertURL API)을 이용하여 등록
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
' 파트너의 연동회원으로 가입된 사업자번호인지 확인합니다.
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
' 파트너의 연동회원으로 회원가입을 요청합니다.
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '아이디, 6자이상 50자 미만
    joinData.id = "userid"
    
    '비밀번호, 6자이상 20자 미만
    joinData.pwd = "pwd_must_be_long_enough"
    
    '파트너링크 아이디
    joinData.LinkID = LinkID
    
    '사업자번호, '-'제외, 10자리
    joinData.CorpNum = "1234567890"
    
    '대표자성명, 최대 100자
    joinData.ceoname = "대표자성명"
    
    '상호명, 최대 200자
    joinData.corpName = "회원상호"
    
    '사업장 주소, 최대 300자
    joinData.addr = "주소"
    
    '업태, 최대 100자
    joinData.bizType = "업태"
    
    '종목, 최대 100자
    joinData.bizClass = "종목"

    '담당자 성명, 최대 100자
    joinData.ContactName = "담당자성명"
    
    '담당자 이메일, 최대 100자
    joinData.ContactEmail = "test@test.com"
    
    '담당자 연락처, 최대 20자
    joinData.ContactTEL = "02-999-9999"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.ContactHP = "010-1234-5678"
    
    '담당자 팩스번호, 최대 20자
    joinData.ContactFAX = "02-999-9998"
    
    
    Set Response = TaxinvoiceService.JoinMember(joinData)
    
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
' 연동회원의 전자세금계산서 API 서비스 과금정보를 확인합니다.
'=========================================================================
Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = TaxinvoiceService.GetChargeInfo(txtCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (발행단가) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (과금유형) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (과금제도) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
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
' 공인인증서 등록 URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================
Private Sub btnGetTaxCertURL_Click()
    Dim url As String
           
    url = TaxinvoiceService.GetTaxCertURL(txtCorpNum.Text, txtUserID.Text)

    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    
    'Internet Explorer Browser 호출
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
        .StatusText = "팝빌 공인인증서 등록 URL"
    End With
    
    Set IE = Nothing
End Sub

'=========================================================================
' 팝빌에 등록된 공인인증서의 유효성을 확인한다.
'=========================================================================
Private Sub btnCheckCertValidation_Click()
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.CheckCertValidation(txtCorpNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 팝빌(www.popbill.com)에 로그인된 팝빌 URL을 반환합니다.
' - 반환된 URL은 보안정책에 따라 30초의 유효시간을 갖습니다.
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim url As String
           
    url = TaxinvoiceService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 인감 및 첨부문서 등록 팝업 URL을 반환합니다.
' - 반환된 URL은 보안정책으로 인해 30초의 유효시간을 갖습니다.
'=========================================================================
Private Sub btnGetSealURL_Click()
    Dim url As String
           
    url = TaxinvoiceService.GetSealURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
   
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 연동회원의 담당자를 신규로 등록합니다.
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디, 6자 이상 50자 미만
    joinData.id = "testkorea"
    
    '비밀번호, 6자 이상 20자 미만
    joinData.pwd = "test@test.com"
    
    '담당자명, 최대 100자
    joinData.personName = "담당자명"
    
    '담당자 연락처, 최대 20자
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.hp = "010-1234-1234"
    
    '담당자 팩스번,최대 20자
    joinData.fax = "070-1234-1234"
    
    '담당자 메일주소, 최대 100자
    joinData.email = "test@test.com"
    
    '회사조회 권한여부, True-회사조회 / False-개인조회
    joinData.searchAllAllowYN = True
    
    '관리자 여부, True-관리자 / False-사용자
    joinData.mgrYN = False
        
    Set Response = TaxinvoiceService.RegistContact(txtCorpNum.Text, joinData)
    
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
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = TaxinvoiceService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(아이디) | personName(성명) | email(이메일) | hp(휴대폰번호) |  fax(팩스번호) | tel(연락처) | " _
         + "regDT(등록일시) | searchAllAllowYN(회사조회 권한여부) | mgrYN(관리자 여부) | state(상태) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchAllAllowYN) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원의 담당자 정보를 수정합니다.
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디
    joinData.id = txtUserID.Text
    
    '담당자 성명, 최대 100자
    joinData.personName = "담당자명_수정"
    
    '담당자 연락처, 최대 20자
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.hp = "010-1234-1234"
        
    '담당자 팩스번호, 최대 20자
    joinData.fax = "070-1234-1234"
    
    '담당자 이메일, 최대 100자
    joinData.email = "test@test.com"

    '회사조회 권한여부, True-회사조회 / False-개인조회
    joinData.searchAllAllowYN = True
    
    '관리자 여부, True-관리자 / False-사용자
    joinData.mgrYN = False
                
    Set Response = TaxinvoiceService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 회사정보를 확인합니다.
'=========================================================================
Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = TaxinvoiceService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname (대표자명) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName (상호) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr (주소) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType (업태) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass (종목) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원의 회사정보를 수정합니다
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '대표자명, 최대 100자
    CorpInfo.ceoname = "대표자"
    
    '상호, 최대 200자
    CorpInfo.corpName = "상호"
    
    '주소, 최대 300자
    CorpInfo.addr = "서울특별시"
    
    '업태, 최대 100자
    CorpInfo.bizType = "업태"
    
    '종목, 최대 100자
    CorpInfo.bizClass = "종목"
    
    Set Response = TaxinvoiceService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
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
    
    MsgBox "연동회원 잔여포인트 : " + CStr(balance)
End Sub

'=========================================================================
' 연동회원 포인트 충전 URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim url As String
           
    url = TaxinvoiceService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 파트너 잔여포인트를 확인합니다.
' - 과금방식이 연동과금인 경우 연동회원 잔여포인트(GetBalance API)
'   를 통해 확인하시기 바랍니다.
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = TaxinvoiceService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "파트너 잔여포인트 : " + CStr(balance)
End Sub

'=========================================================================
' 파트너 포인트 충전 URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
           
    url = TaxinvoiceService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 세금계산서 관리번호 중복여부를 확인합니다.
' - 관리번호는 1~24자리로 숫자, 영문 '-', '_' 조합으로 구성할 수 있습니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set Response = TaxinvoiceService.checkMgtKeyInUse(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 전자세금계산서 유통사업자의 메일 목록을 확인합니다.
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
    
    tmp = "유통사업자 이메일 목록" + vbCrLf
    For Each email In resultList
        tmp = tmp + email + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 1건의 세금계산서를 즉시발행 처리합니다.
' - 세금계산서 항목별 정보는 "[전자세금계산서 API 연동매뉴얼] > 4.1. (세금)계산서
'   구성"을 참조하시기 바랍니다.
' - [발행완료] 세금계산서는 연동회원의 국세청 전송설정에 따라
'   익일전송/즉시전송 처리됩니다. 기본설정(익일전송)
' - 국세청 전송설정은 "팝빌 로그인" > [전자세금계산서] > [환경설정] >
'   [전자세금계산서 관리] > [국세청 전송 및 지연발행 설정] 탭에서
'   확인할 수 있습니다.
' - 국세청 전송정책에 대한 사항은 "[전자세금계산서 API 연동매뉴얼] >
'   1.3. 국세청 전송 정책" 을 참조하시기 바랍니다.
'=========================================================================
Private Sub btnRegistIssue_Click()
    Dim Taxinvoice As New PBTaxinvoice
    
    '[필수] 작성일자, 표시형식 (yyyyMMdd) ex)20190207
    Taxinvoice.writeDate = "20190207"
    
    '[필수] 발행형태, [정발행, 역발행, 위수탁] 중 기재
    Taxinvoice.issueType = "정발행"
    
    '[필수] {정과금, 역과금} 중 기재, '역과금'은 역발행 프로세스에서만 이용가능
    '- 정과금(공급자 과금), 역과금(공급받는자 과금)
    Taxinvoice.chargeDirection = "정과금"
    
    '[필수] 영수/청구, [영수, 청구] 중 기재
    Taxinvoice.purposeType = "영수"
    
    '[필수] 발행시점
    Taxinvoice.issueTiming = "직접발행"
    
    '[필수] 과세형태, [과세, 영세, 면세] 중 기재
    Taxinvoice.taxType = "과세"
    
    
    '=========================================================================
    '                              공급자 정보
    '=========================================================================
        
    '[필수] 공급자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoicerCorpNum = txtCorpNum.Text
    
    '[필수] 공급자 종사업장 식별번호. 필요시 숫자 4자리 기재
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
    
    ' 발행시 알림문자 전송여부 (정발행에서만 사용가능)
    ' - 공급받는자 주)담당자 휴대폰번호(invoiceeHP1)로 전송
    ' - 전송시 포인트가 차감되며 전송실패하는 경우 포인트 환불처리
    Taxinvoice.invoicerSMSSendYN = False
    
    
    '=========================================================================
    '                            공급받는자 정보
    '=========================================================================
        
    '[필수] 공급받는자 구분, [사업자, 개인, 외국인] 중 기재
    Taxinvoice.invoiceeType = "사업자"
    
    '[필수] 공급받는자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoiceeCorpNum = "8888888888"
    
    '[필수] 공급받는자 종사업장 식별번호. 필요시 숫자 4자리 기재
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
    
    
    '=========================================================================
    '                            세금계산서 정보
    '=========================================================================
    
    '[필수] 공급가액 합계
    Taxinvoice.supplyCostTotal = "200000"
    
    '[필수] 세액 합계
    Taxinvoice.taxTotal = "20000"
    
    '[필수] 합계금액, 공급가액 합계 + 세액합계
    Taxinvoice.totalAmount = "220000"
    
    '기재 상 '일련번호' 항목
    Taxinvoice.serialNum = "123"
    
    '기재 상 '권' 항목, 최대값 32767
    '미기재시 Taxinvoice.kwon = ""
    Taxinvoice.kwon = "1"
    
    '기재 상 '호' 항목, 최대값 32767
    '미기재시 Taxinvoice.kwon = ""
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
    '             상세항목(품목) 정보 > 배열로 99개까지 기재 가능
    '=========================================================================
    
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20190207"   '거래일자  yyyyMMdd
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
    newDetail2.purchaseDT = "20190207"   '거래일자  yyyyMMdd
    newDetail2.itemName = "품명2"        '품목명
    newDetail2.spec = "규격"             '규격
    newDetail2.qty = "1"                 '수량
    newDetail2.unitCost = "100000"       '단가
    newDetail2.supplyCost = "100000"     '공급가액
    newDetail2.tax = "10000"             '세액
    newDetail2.remark = "비고"           '비고
    
    Taxinvoice.detailList.Add newDetail2
    
    
    '=========================================================================
    '             추가담당자 정보 > 배열로 5개까지 기재 가능
    ' - 세금계산서 발행안내 메일을 수신받을 공급받는자 담당자가 다수인 경우
    ' 담당자 정보를 추가하여 발행안내메일을 다수에게 전송할 수 있습니다.
    '=========================================================================
    
    Set Taxinvoice.addContactList = New Collection
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                 '일련번호, 1부터 순차기재
    newContact.ContactName = "담당자 성명"   '담당자명
    newContact.email = "test2@test.com"      '담당자 메일주소
    Taxinvoice.addContactList.Add newContact
    
    Dim newContact2 As New PBTIContact
    newContact2.serialNum = 2                '일련번호, 1부터 순차기재
    newContact2.ContactName = "담당자 성명"  '담당자명
    newContact2.email = "test2@test.com"     '담당자 메일주소
    Taxinvoice.addContactList.Add newContact2
        
    
    '거래명세서 동시작성 여부
    Taxinvoice.writeSpecification = False
    
    '거래명세서 동시작성시 거래명세서 관리번호, 미기재시 세금계산서 관리번호로 자동작성
    Taxinvoice.dealInvoiceMgtKey = ""
    
    '지연발행 강제여부(forceIssue)
    '발행마감일이 지난 세금계산서를 발행하는 경우, 가산세가 부과될 수 있습니다.
    '가산세가 부과되더라도 발행을 해야하는 경우에는 forceIssue의 값을
    'true로 선언하여 발행(Issue API)를 호출하시면 됩니다.
    Taxinvoice.forceIssue = False
    
    '메모
    Taxinvoice.memo = ""
    
    '발행안내 메일제목, 공백처리시 기본제목으로 전송
    Taxinvoice.emailSubject = ""
    
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.RegistIssue(txtCorpNum.Text, Taxinvoice)
    
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '메모
    memo = "발행 취소 메모"
    
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 1건의 전자세금계산서를 삭제합니다.
' - 세금계산서를 삭제해야만 문서관리번호(mgtKey)를 재사용할 수 있습니다.
' - 삭제가능한 문서 상태 : 임시저장, 발행취소, 역)발행 거부/취소
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
        
    Set Response = TaxinvoiceService.Delete(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
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
    Dim writeSpecification As Boolean
    Dim Taxinvoice As New PBTaxinvoice
    
    '[필수] 작성일자, 표시형식 (yyyyMMdd) ex)20190207
    Taxinvoice.writeDate = "20190207"
    
    '[필수] 발행형태, [정발행, 역발행, 위수탁] 중 기재
    Taxinvoice.issueType = "정발행"
    
    '[필수] {정과금, 역과금} 중 기재, '역과금'은 역발행 프로세스에서만 이용가능
    '- 정과금(공급자 과금), 역과금(공급받는자 과금)
    Taxinvoice.chargeDirection = "정과금"
    
    '[필수] 영수/청구, [영수, 청구] 중 기재
    Taxinvoice.purposeType = "영수"
    
    '[필수] 발행시점
    Taxinvoice.issueTiming = "직접발행"
    
    '[필수] 과세형태, [과세, 영세, 면세] 중 기재
    Taxinvoice.taxType = "과세"
    
    
    '=========================================================================
    '                              공급자 정보
    '=========================================================================
        
    '[필수] 공급자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoicerCorpNum = txtCorpNum.Text
    
    '[필수] 공급자 종사업장 식별번호. 필요시 숫자 4자리 기재
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
    Taxinvoice.invoicerTEL = "070-4304-2991"
    
    '공급자 담당자 휴대폰번호
    Taxinvoice.invoicerHP = "010-000-2222"
    
    ' 발행시 알림문자 전송여부 (정발행에서만 사용가능)
    ' - 공급받는자 주)담당자 휴대폰번호(invoiceeHP1)로 전송
    ' - 전송시 포인트가 차감되며 전송실패하는 경우 포인트 환불처리
    Taxinvoice.invoicerSMSSendYN = False
    
    
    '=========================================================================
    '                            공급받는자 정보
    '=========================================================================
        
    '[필수] 공급받는자 구분, [사업자, 개인, 외국인] 중 기재
    Taxinvoice.invoiceeType = "사업자"
    
    '[필수] 공급받는자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoiceeCorpNum = "8888888888"
    
    '[필수] 공급받는자 종사업장 식별번호. 필요시 숫자 4자리 기재
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
            
    
    '=========================================================================
    '                            세금계산서 정보
    '=========================================================================
    
    '[필수] 공급가액 합계
    Taxinvoice.supplyCostTotal = "200000"
    
    '[필수] 세액 합계
    Taxinvoice.taxTotal = "20000"
    
    '[필수] 합계금액, 공급가액 합계 + 세액합계
    Taxinvoice.totalAmount = "220000"
    
    '기재 상 '일련번호' 항목
    Taxinvoice.serialNum = "123"
    
    '기재 상 '권' 항목, 최대값 32767
    ' 미기재시 Taxinvoice.kwon = ""
    Taxinvoice.kwon = "1"
    
    '기재 상 '호' 항목, 최대값 32767
    ' 미기재시 Taxinvoice.kwon = ""
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
    '         수정세금계산서 정보 (수정세금계산서 작성시에만 기재)
    ' - 수정세금계산서 관련 정보는 연동매뉴얼 또는 개발가이드 링크 참조
    ' - [참고] 수정세금계산서 작성방법 안내 - http://blog.linkhub.co.kr/650
    '========================================================================='
    
    ' 수정사유코드, 수정사유에 따라 1~6중 선택기재
    Taxinvoice.modifyCode = ""
    
    ' 원본세금계산서의 ItemKey, 문서확인 (GetInfo API)의 응답결과(ItemKey 항목) 확인
    Taxinvoice.originalTaxinvoiceKey = ""
        
    
    '=========================================================================
    '             상세항목(품목) 정보 > 배열로 99개까지 기재 가능
    '=========================================================================
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20190207"   '거래일자  yyyyMMdd
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
    newDetail2.purchaseDT = "20190207"   '거래일자  yyyyMMdd
    newDetail2.itemName = "품명2"        '품목명
    newDetail2.spec = "규격"             '규격
    newDetail2.qty = "1"                 '수량
    newDetail2.unitCost = "100000"       '단가
    newDetail2.supplyCost = "100000"     '공급가액
    newDetail2.tax = "10000"             '세액
    newDetail2.remark = "비고"           '비고
    
    Taxinvoice.detailList.Add newDetail2
    
    
    '=========================================================================
    '             추가담당자 정보 > 배열로 5개까지 기재 가능
    ' - 세금계산서 발행안내 메일을 수신받을 공급받는자 담당자가 다수인 경우
    ' 담당자 정보를 추가하여 발행안내메일을 다수에게 전송할 수 있습니다.
    '=========================================================================
    Set Taxinvoice.addContactList = New Collection
    
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                 '일련번호, 1부터 순차기재
    newContact.ContactName = "담당자 성명"   '담당자명
    newContact.email = "test2@test.com"      '담당자 메일주소
    Taxinvoice.addContactList.Add newContact
    
    Dim newContact2 As New PBTIContact
    newContact2.serialNum = 2                '일련번호, 1부터 순차기재
    newContact2.ContactName = "담당자 성명"  '담당자명
    newContact2.email = "test2@test.com"     '담당자 메일주소
    Taxinvoice.addContactList.Add newContact2
    
    '거래명세서 동시작성 여부
    writeSpecification = False
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.Register(txtCorpNum.Text, Taxinvoice, writeSpecification)
    
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
    
    '[필수] 작성일자, 표시형식 (yyyyMMdd) ex)20190207
    Taxinvoice.writeDate = "20190207"
    
    '[필수] 발행형태, [정발행, 역발행, 위수탁] 중 기재
    Taxinvoice.issueType = "정발행"
    
    '[필수] {정과금, 역과금} 중 기재, '역과금'은 역발행 프로세스에서만 이용가능
    '- 정과금(공급자 과금), 역과금(공급받는자 과금)
    Taxinvoice.chargeDirection = "정과금"
    
    '[필수] 영수/청구, [영수, 청구] 중 기재
    Taxinvoice.purposeType = "영수"
    
    '[필수] 발행시점
    Taxinvoice.issueTiming = "직접발행"
    
    '[필수] 과세형태, [과세, 영세, 면세] 중 기재
    Taxinvoice.taxType = "과세"
    
    
    '=========================================================================
    '                              공급자 정보
    '=========================================================================
        
    '[필수] 공급자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoicerCorpNum = txtCorpNum.Text
    
    '[필수] 공급자 종사업장 식별번호. 필요시 숫자 4자리 기재
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
    
    ' 발행시 알림문자 전송여부 (정발행에서만 사용가능)
    ' - 공급받는자 주)담당자 휴대폰번호(invoiceeHP1)로 전송
    ' - 전송시 포인트가 차감되며 전송실패하는 경우 포인트 환불처리
    Taxinvoice.invoicerSMSSendYN = False
    
    
    '=========================================================================
    '                            공급받는자 정보
    '=========================================================================
        
    '[필수] 공급받는자 구분, [사업자, 개인, 외국인] 중 기재
    Taxinvoice.invoiceeType = "사업자"
    
    '[필수] 공급받는자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoiceeCorpNum = "8888888888"
    
    '[필수] 공급받는자 종사업장 식별번호. 필요시 숫자 4자리 기재
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
    
    '=========================================================================
    '                            세금계산서 정보
    '=========================================================================
    
    '[필수] 공급가액 합계
    Taxinvoice.supplyCostTotal = "200000"
    
    '[필수] 세액 합계
    Taxinvoice.taxTotal = "20000"
    
    '[필수] 합계금액, 공급가액 합계 + 세액합계
    Taxinvoice.totalAmount = "220000"
    
    '기재 상 '일련번호' 항목
    Taxinvoice.serialNum = "123"
    
    '기재 상 '권' 항목, 최대값 32767
    '미기재시 Taxinvoice.kwon = ""
    Taxinvoice.kwon = "1"
    
    '기재 상 '호' 항목, 최대값 32767
    '미기재시 Taxinvoice.kwon = ""
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
    '             상세항목(품목) 정보 > 배열로 99개까지 기재 가능
    '=========================================================================
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20190207"   '거래일자  yyyyMMdd
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
    newDetail2.purchaseDT = "20190207"   '거래일자  yyyyMMdd
    newDetail2.itemName = "품명2"        '품목명
    newDetail2.spec = "규격"             '규격
    newDetail2.qty = "1"                 '수량
    newDetail2.unitCost = "100000"       '단가
    newDetail2.supplyCost = "100000"     '공급가액
    newDetail2.tax = "10000"             '세액
    newDetail2.remark = "비고"           '비고
    
    Taxinvoice.detailList.Add newDetail2
    
    
    '=========================================================================
    '             추가담당자 정보 > 배열로 5개까지 기재 가능
    ' - 세금계산서 발행안내 메일을 수신받을 공급받는자 담당자가 다수인 경우
    ' 담당자 정보를 추가하여 발행안내메일을 다수에게 전송할 수 있습니다.
    '=========================================================================
    Set Taxinvoice.addContactList = New Collection
    
    Dim newContact As New PBTIContact
    newContact.serialNum = 1                 '일련번호, 1부터 순차기재
    newContact.ContactName = "담당자 성명"   '담당자명
    newContact.email = "test2@test.com"      '담당자 메일주소
    Taxinvoice.addContactList.Add newContact
    
    Dim newContact2 As New PBTIContact
    newContact2.serialNum = 2                '일련번호, 1부터 순차기재
    newContact2.ContactName = "담당자 성명"  '담당자명
    newContact2.email = "test2@test.com"     '담당자 메일주소
    
    Taxinvoice.addContactList.Add newContact2
    
    '거래명세서 동시작성여부
    writeSpecification = False
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.Update(txtCorpNum.Text, KeyType, txtMgtKey.Text, Taxinvoice, writeSpecification)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'======================================================================================================================
' [임시저장] 상태의 세금계산서를 [발행]처리 합니다.
' - 발행(Issue API)를 호출하는 시점에서 포인트가 차감됩니다.
' - [발행완료] 세금계산서는 연동회원의 국세청 전송설정에 따라 익일/즉시전송 처리됩니다. 기본설정(익일전송)
' - 국세청 전송설정은 "팝빌 로그인" > [전자세금계산서] > [환경설정] >
'   [전자세금계산서 관리] > [국세청 전송 및 지연발행 설정] 탭에서 확인할 수 있습니다.
' - 국세청 전송정책에 대한 사항은 "[전자세금계산서 API 연동매뉴얼] > 1.3. 국세청 전송 정책" 을 참조하시기 바랍니다
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
        
    
    Set Response = TaxinvoiceService.Issue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, emailSubject, forceIssue)
    
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
    memo = "발행취소 메모"
    
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 1건의 전자세금계산서를 삭제합니다.
' - 세금계산서를 삭제해야만 문서관리번호(mgtKey)를 재사용할 수 있습니다.
' - 삭제가능한 문서 상태 : 임시저장, 발행취소, 역)발행 거부/취소
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
    
    Set Response = TaxinvoiceService.Delete(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
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
' - 국세청 전송에 관한 사항은 "[전자세금계산서 API 연동매뉴얼] > 1.3 국세청
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
    
    Set Response = TaxinvoiceService.SendToNTS(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================================================================================
'[공급받는자]가 공급자에게 1건의 역발행 세금계산서를 [즉시 요청]합니다.
' - 세금계산서 항목별 정보는 "[전자세금계산서 API 연동매뉴얼] > 4.1. (세금)계산서구성"을 참조하시기 바랍니다.
' - 역발행 세금계산서 프로세스를 구현하기 위해서는 공급자/공급받는자가 모두 팝빌에 회원이여야 합니다.
' - 역발행 즉시요청후 공급자가 [발행] 처리시 포인트가 차감되며 역발행 세금계산서 항목중 과금방향(ChargeDirection)에 기재한 값에 따라
'   정과금(공급자과금) 또는 역과금(공급받는자과금) 처리됩니다.
'=========================================================================================================================================
Private Sub btnRegistRequest_Click()
    Dim Taxinvoice As New PBTaxinvoice
    
    '[필수] 작성일자, 표시형식 (yyyyMMdd) ex)20190207
    Taxinvoice.writeDate = "20190207"
    
    '[필수] 발행형태, [정발행, 역발행, 위수탁] 중 기재
    Taxinvoice.issueType = "역발행"
    
    '[필수] {정과금, 역과금} 중 기재, '역과금'은 역발행 프로세스에서만 이용가능
    '- 정과금(공급자 과금), 역과금(공급받는자 과금)
    Taxinvoice.chargeDirection = "정과금"
    
    '[필수] 영수/청구, [영수, 청구] 중 기재
    Taxinvoice.purposeType = "영수"
    
    '[필수] 발행시점
    Taxinvoice.issueTiming = "직접발행"
    
    '[필수] 과세형태, [과세, 영세, 면세] 중 기재
    Taxinvoice.taxType = "과세"
    
    
    '=========================================================================
    '                              공급자 정보
    '=========================================================================
        
    '[필수] 공급자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoicerCorpNum = "8888888888"
    
    '[필수] 공급자 종사업장 식별번호. 필요시 숫자 4자리 기재
    Taxinvoice.invoicerTaxRegID = ""
    
    '[필수] 공급자 상호
    Taxinvoice.invoicerCorpName = "공급자 상호"
    
    '[정발행시 필수] 공급자 문서관리번호, 1~24자리 (숫자, 영문, '-', '_') 조합으로
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
    
    
    '=========================================================================
    '                            공급받는자 정보
    '=========================================================================
        
    '[필수] 공급받는자 구분, [사업자, 개인, 외국인] 중 기재
    Taxinvoice.invoiceeType = "사업자"
    
    '[필수] 공급받는자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoiceeCorpNum = txtCorpNum.Text
    
    '[필수] 공급받는자 종사업장 식별번호. 필요시 숫자 4자리 기재
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
    
    ' 역발행 요청시 알림문자 전송여부 (역발행에서만 사용가능)
    ' - 공급자 담당자 휴대폰번호(invoicerHP)로 전송
    ' - 전송시 포인트가 차감되며 전송실패하는 경우 포인트 환불처리
    Taxinvoice.invoiceeSMSSendYN = False
    
    '=========================================================================
    '                            세금계산서 정보
    '=========================================================================
    
    '[필수] 공급가액 합계
    Taxinvoice.supplyCostTotal = "200000"
    
    '[필수] 세액 합계
    Taxinvoice.taxTotal = "20000"
    
    '[필수] 합계금액, 공급가액 합계 + 세액합계
    Taxinvoice.totalAmount = "220000"
    
    '기재 상 '일련번호' 항목
    Taxinvoice.serialNum = "123"
    
    '기재 상 '권' 항목, 최대값 32767
    '미기재시 Taxinvoice.kwon = ""
    Taxinvoice.kwon = "1"
    
    '기재 상 '호' 항목, 최대값 32767
    '미기재시 Taxinvoice.kwon = ""
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
    '             상세항목(품목) 정보 > 배열로 99개까지 기재 가능
    '=========================================================================
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20190207"   '거래일자  yyyyMMdd
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
    newDetail2.purchaseDT = "20190207"   '거래일자  yyyyMMdd
    newDetail2.itemName = "품명2"        '품목명
    newDetail2.spec = "규격"             '규격
    newDetail2.qty = "1"                 '수량
    newDetail2.unitCost = "100000"       '단가
    newDetail2.supplyCost = "100000"     '공급가액
    newDetail2.tax = "10000"             '세액
    newDetail2.remark = "비고"           '비고
    
    Taxinvoice.detailList.Add newDetail2
        
    '메모
    Taxinvoice.memo = "즉시요청 메모"
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.RegistRequest(txtCorpNum.Text, Taxinvoice)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'======================================================================================================================
' [역발행대기] 상태의 세금계산서를 [발행]처리 합니다.
' - 발행(Issue API)를 호출하는 시점에서 포인트가 차감됩니다.
' - [발행완료] 세금계산서는 연동회원의 국세청 전송설정에 따라 익일/즉시전송 처리됩니다. 기본설정(익일전송)
' - 국세청 전송설정은 "팝빌 로그인" > [전자세금계산서] > [환경설정] >
'   [전자세금계산서 관리] > [국세청 전송 및 지연발행 설정] 탭에서 확인할 수 있습니다.
' - 국세청 전송정책에 대한 사항은 "[전자세금계산서 API 연동매뉴얼] > 1.3. 국세청 전송 정책" 을 참조하시기 바랍니다
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
    
    Set Response = TaxinvoiceService.Issue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, emailSubject, forceIssue)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 공급받는자에게 요청받은 역발행 세금계산서를 [거부]처리 합니다.
' - 세금계산서의 문서관리번호를 재사용하기 위해서는 삭제 (Delete API) 를
'   호출하여 [삭제] 처리해야 합니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '메모
    memo = "역)발행 요청 거부 메모"
    
    Set Response = TaxinvoiceService.Refuse(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '메모
    memo = "발행취소 메모"
    
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 역발행 세금계산서를 [요청취소] 처리합니다.
' - [취소]한 세금계산서의 문서관리번호를 재사용하기 위해서는 삭제 (Delete API) 를 호출해야 합니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '메모
    memo = "역)발행 요청 취소 메모"
    
    Set Response = TaxinvoiceService.CancelRequest(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 1건의 전자세금계산서를 삭제합니다.
' - 세금계산서를 삭제해야만 문서관리번호(mgtKey)를 재사용할 수 있습니다.
' - 삭제가능한 문서 상태 : 임시저장, 발행취소, 역)발행 거부/취소
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select

    Set Response = TaxinvoiceService.Delete(txtCorpNum.Text, KeyType, txtMgtKey.Text)

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
    
    '[필수] 작성일자, 표시형식 (yyyyMMdd) ex)20190207
    Taxinvoice.writeDate = "20190207"
    
    '[필수] 발행형태, [정발행, 역발행, 위수탁] 중 기재
    Taxinvoice.issueType = "역발행"
    
    '[필수] {정과금, 역과금} 중 기재, '역과금'은 역발행 프로세스에서만 이용가능
    '- 정과금(공급자 과금), 역과금(공급받는자 과금)
    Taxinvoice.chargeDirection = "정과금"
    
    '[필수] 영수/청구, [영수, 청구] 중 기재
    Taxinvoice.purposeType = "영수"
    
    '[필수] 발행시점
    Taxinvoice.issueTiming = "직접발행"
    
    '[필수] 과세형태, [과세, 영세, 면세] 중 기재
    Taxinvoice.taxType = "과세"
    
    
    '=========================================================================
    '                              공급자 정보
    '=========================================================================
        
    '[필수] 공급자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoicerCorpNum = "8888888888"
    
    '[필수] 공급자 종사업장 식별번호. 필요시 숫자 4자리 기재
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
    Taxinvoice.invoiceeCorpNum = txtCorpNum.Text
    
    '[필수] 공급받는자 종사업장 식별번호. 필요시 숫자 4자리 기재
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
    
    ' 역발행 요청시 알림문자 전송여부 (역발행에서만 사용가능)
    ' - 공급자 담당자 휴대폰번호(invoicerHP)로 전송
    ' - 전송시 포인트가 차감되며 전송실패하는 경우 포인트 환불처리
    Taxinvoice.invoiceeSMSSendYN = False
            
    
    '=========================================================================
    '                            세금계산서 정보
    '=========================================================================
    
    '[필수] 공급가액 합계
    Taxinvoice.supplyCostTotal = "200000"
    
    '[필수] 세액 합계
    Taxinvoice.taxTotal = "20000"
    
    '[필수] 합계금액, 공급가액 합계 + 세액합계
    Taxinvoice.totalAmount = "220000"
    
    '기재 상 '일련번호' 항목
    Taxinvoice.serialNum = "123"
    
    '기재 상 '권' 항목, 최대값 32767
    '미기재시 Taxinvoice.kwon = ""
    Taxinvoice.kwon = "1"
    
    '기재 상 '호' 항목, 최대값 32767
    '미기재시 Taxinvoice.kwon = ""
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
    '             상세항목(품목) 정보 > 배열로 99개까지 기재 가능
    '=========================================================================
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20190207"   '거래일자  yyyyMMdd
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
    newDetail2.purchaseDT = "20190207"   '거래일자  yyyyMMdd
    newDetail2.itemName = "품명2"        '품목명
    newDetail2.spec = "규격"             '규격
    newDetail2.qty = "1"                 '수량
    newDetail2.unitCost = "100000"       '단가
    newDetail2.supplyCost = "100000"     '공급가액
    newDetail2.tax = "10000"             '세액
    newDetail2.remark = "비고"           '비고
    
    Taxinvoice.detailList.Add newDetail2

    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.Register(txtCorpNum.Text, Taxinvoice, False)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' [임시저장] 상태의  역)발행 세금계산서의 항목을 수정합니다.
' - 세금계산서 항목별 정보는 "[전자세금계산서 API 연동매뉴얼] > 4.1. (세금)계산서
'   구성"을 참조하시기 바랍니다.
'=========================================================================
Private Sub btnUpdate_rev_Click()
    Dim KeyType As MgtKeyType
    
    '세금계산서 발행유형, SELL-매출, BUY-매입, TRUSTEE-위수탁
    KeyType = BUY
    
    Dim Taxinvoice As New PBTaxinvoice
    
    '[필수] 작성일자, 표시형식 (yyyyMMdd) ex)20190207
    Taxinvoice.writeDate = "20190207"
    
    '[필수] 발행형태, [정발행, 역발행, 위수탁] 중 기재
    Taxinvoice.issueType = "역발행"
    
    '[필수] {정과금, 역과금} 중 기재, '역과금'은 역발행 프로세스에서만 이용가능
    '- 정과금(공급자 과금), 역과금(공급받는자 과금)
    Taxinvoice.chargeDirection = "정과금"
    
    '[필수] 영수/청구, [영수, 청구] 중 기재
    Taxinvoice.purposeType = "영수"
    
    '[필수] 발행시점
    Taxinvoice.issueTiming = "직접발행"
    
    '[필수] 과세형태, [과세, 영세, 면세] 중 기재
    Taxinvoice.taxType = "과세"
    
    
    '=========================================================================
    '                              공급자 정보
    '=========================================================================
        
    '[필수] 공급자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoicerCorpNum = "8888888888"
    
    '[필수] 공급자 종사업장 식별번호. 필요시 숫자 4자리 기재
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
    
    '=========================================================================
    '                            공급받는자 정보
    '=========================================================================
        
    '[필수] 공급받는자 구분, [사업자, 개인, 외국인] 중 기재
    Taxinvoice.invoiceeType = "사업자"
    
    '[필수] 공급받는자 사업자번호, '-' 제외 10자리
    Taxinvoice.invoiceeCorpNum = "1234567890"
    
    '[필수] 공급받는자 종사업장 식별번호. 필요시 숫자 4자리 기재
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
    
    ' 역발행 요청시 알림문자 전송여부 (역발행에서만 사용가능)
    ' - 공급자 담당자 휴대폰번호(invoicerHP)로 전송
    ' - 전송시 포인트가 차감되며 전송실패하는 경우 포인트 환불처리
    Taxinvoice.invoiceeSMSSendYN = False
            
    
    '=========================================================================
    '                            세금계산서 정보
    '=========================================================================
    
    '[필수] 공급가액 합계
    Taxinvoice.supplyCostTotal = "200000"
    
    '[필수] 세액 합계
    Taxinvoice.taxTotal = "20000"
    
    '[필수] 합계금액, 공급가액 합계 + 세액합계
    Taxinvoice.totalAmount = "220000"
    
    '기재 상 '일련번호' 항목
    Taxinvoice.serialNum = "123"
    
    '기재 상 '권' 항목, 최대값 32767
    '미기재시 Taxinvoice.kwon = ""
    Taxinvoice.kwon = "1"
    
    '기재 상 '호' 항목, 최대값 32767
    '미기재시 Taxinvoice.kwon = ""
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
    '             상세항목(품목) 정보 > 배열로 99개까지 기재 가능
    '=========================================================================
    Set Taxinvoice.detailList = New Collection
    
    Dim newDetail As New PBTIDetail
    
    newDetail.serialNum = 1             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20190207"   '거래일자  yyyyMMdd
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
    newDetail2.purchaseDT = "20190207"   '거래일자  yyyyMMdd
    newDetail2.itemName = "품명2"        '품목명
    newDetail2.spec = "규격"             '규격
    newDetail2.qty = "1"                 '수량
    newDetail2.unitCost = "100000"       '단가
    newDetail2.supplyCost = "100000"     '공급가액
    newDetail2.tax = "10000"             '세액
    newDetail2.remark = "비고"           '비고
    
    Taxinvoice.detailList.Add newDetail2
    
    Dim Response As PBResponse
    
    Set Response = TaxinvoiceService.Update(txtCorpNum.Text, KeyType, txtMgtKey.Text, Taxinvoice, False)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 공급받는자가 공급자에게 1건의 역발행 세금계산서를 발행요청 합니다.
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
    
    
    Set Response = TaxinvoiceService.Request(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'======================================================================================================================
' [역발행대기] 상태의 세금계산서를 [발행]처리 합니다.
' - 발행(Issue API)를 호출하는 시점에서 포인트가 차감됩니다.
' - [발행완료] 세금계산서는 연동회원의 국세청 전송설정에 따라 익일/즉시전송 처리됩니다. 기본설정(익일전송)
' - 국세청 전송설정은 "팝빌 로그인" > [전자세금계산서] > [환경설정] >
'   [전자세금계산서 관리] > [국세청 전송 및 지연발행 설정] 탭에서 확인할 수 있습니다.
' - 국세청 전송정책에 대한 사항은 "[전자세금계산서 API 연동매뉴얼] > 1.3. 국세청 전송 정책" 을 참조하시기 바랍니다
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '메모
    memo = "역발행 세금계산서 발행"
    
    '공급받는자에게 전송되는 발행안내메일 제목, 미기재시 기본양식으로 전송
    emailSubject = ""
    
    '지연발행 강제여부, 기본값 - False
    '발행마감일이 지난 세금계산서를 발행하는 경우, 가산세가 부과될 수 있습니다.
    '지연발행 세금계산서를 신고해야 하는 경우 forceIssue 값을 True로 선언하여 발행(Issue API)을 호출할 수 있습니다.
    forceIssue = False
    
    Set Response = TaxinvoiceService.Issue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo, emailSubject, forceIssue)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
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
    
    Set Response = TaxinvoiceService.Refuse(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
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
    memo = "발행취소 메모"
    
    Set Response = TaxinvoiceService.CancelIssue(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 역발행 세금계산서를 [요청취소] 처리합니다.
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
    memo = "역)발행 요청 취소 메모"
    
    Set Response = TaxinvoiceService.CancelRequest(txtCorpNum.Text, KeyType, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 1건의 전자세금계산서를 삭제합니다.
' - 세금계산서를 삭제해야만 문서관리번호(mgtKey)를 재사용할 수 있습니다.
' - 삭제가능한 문서 상태 : 임시저장, 발행취소, 역)발행 거부/취소
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

    Set Response = TaxinvoiceService.Delete(txtCorpNum.Text, KeyType, txtMgtKey.Text)

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

    Set Response = TaxinvoiceService.AttachFile(txtCorpNum.Text, KeyType, txtMgtKey.Text, FilePath)

    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If

    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 세금계산서에 첨부된 파일의 목록을 확인합니다.
' - 응답항목 중 파일아이디(AttachedFile) 항목은 파일삭제(DeleteFile API)
'   호출시 이용할 수 있습니다.
'=========================================================================
Private Sub btnGetFiles_Click()
    Dim resultList As Collection
    Dim KeyType As MgtKeyType
    Dim tmp As String
    Dim file As PBAttachFile
    
    Set resultList = TaxinvoiceService.GetFiles(txtCorpNum.Text, KeyType, txtMgtKey.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "serialNum(일련번호) | attachedfile(파일아이디) | displayName(첨부파일명) |  RegDT(첨부일시)" + vbCrLf
    
    For Each file In resultList
        tmp = tmp + CStr(file.serialNum) + " | " + file.AttachedFile + " | " + file.DisplayName + " | " + file.regDT + vbCrLf
        txtFileID.Text = file.AttachedFile
        
    Next
    
    MsgBox tmp
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
    
    Set Response = TaxinvoiceService.DeleteFile(txtCorpNum.Text, KeyType, txtMgtKey.Text, txtFileID.Text)
            
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
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
    Dim tmp As String
   
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
    
    Set tiInfo = TaxinvoiceService.GetInfo(txtCorpNum.Text, KeyType, txtMgtKey.Text)
     
    If tiInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "itemKey (팝빌 관리번호) : " + tiInfo.itemKey + vbCrLf
    tmp = tmp + "taxType (과세형태) : " + tiInfo.taxType + vbCrLf
    tmp = tmp + "writeDate (작성일자) : " + tiInfo.writeDate + vbCrLf
    tmp = tmp + "regDT (임시저장 일자) : " + tiInfo.regDT + vbCrLf
    tmp = tmp + "issueType (발행형태) : " + tiInfo.issueType + vbCrLf
    tmp = tmp + "supplyCostTotal (공급가액 합계) : " + tiInfo.supplyCostTotal + vbCrLf
    tmp = tmp + "taxTotal (세액 합계) : " + tiInfo.taxTotal + vbCrLf
    tmp = tmp + "purposeType (영수/청구) : " + tiInfo.purposeType + vbCrLf
    tmp = tmp + "issueDT (발행일시) : " + tiInfo.issueDT + vbCrLf
    tmp = tmp + "stateDT (상태 변경일시) : " + tiInfo.stateDT + vbCrLf
    tmp = tmp + "lateIssueYN (지연발행 여부) : " + CStr(tiInfo.lateIssueYN) + vbCrLf
    tmp = tmp + "openYN (개봉여부) : " + CStr(tiInfo.openYN) + vbCrLf
    tmp = tmp + "openDT (개봉일시) : " + tiInfo.openDT + vbCrLf
    tmp = tmp + "stateCode (상태코드) : " + CStr(tiInfo.stateCode) + vbCrLf
    tmp = tmp + "stateMemo (상태메모) : " + tiInfo.stateMemo + vbCrLf
    tmp = tmp + "ntsresult (국세청 전송결과) : " + tiInfo.ntsresult + vbCrLf
    tmp = tmp + "ntsconfirmNum (국세청승인번호) : " + tiInfo.ntsconfirmNum + vbCrLf
    tmp = tmp + "ntssendDT (국세청 전송일시) : " + tiInfo.ntssendDT + vbCrLf
    tmp = tmp + "ntsresultDT (국세청 결과 수신일시) : " + tiInfo.ntsresultDT + vbCrLf
    tmp = tmp + "ntssendErrCode (전송실패 사유코드) : " + tiInfo.ntssendErrCode + vbCrLf
    tmp = tmp + "modifyCode (수정 사유코드) : " + tiInfo.modifyCode + vbCrLf
    tmp = tmp + "interOPYN (연동문서 여부) : " + CStr(tiInfo.interOPYN) + vbCrLf
    tmp = tmp + "invoicerCorpName (공급자 상호) : " + tiInfo.invoicerCorpName + vbCrLf
    tmp = tmp + "invoicerCorpNum (공급자 사업자번호) : " + tiInfo.invoicerCorpNum + vbCrLf
    tmp = tmp + "invoicerMgtKey (공급자 문서관리번호) : " + tiInfo.invoicerMgtKey + vbCrLf
    tmp = tmp + "invoicerPrintYN (공급자 인쇄여부) : " + CStr(tiInfo.invoicerPrintYN) + vbCrLf
    tmp = tmp + "invoiceeCorpName (공급받는자 상호) : " + tiInfo.invoiceeCorpName + vbCrLf
    tmp = tmp + "invoiceeCorpNum (공급받는자 사업자번호) : " + tiInfo.invoiceeCorpNum + vbCrLf
    tmp = tmp + "invoiceeMgtKey (공급받는자 문서관리번호) : " + tiInfo.invoiceeMgtKey + vbCrLf
    tmp = tmp + "invoiceePrintYN (공급받는자 인쇄여부) : " + CStr(tiInfo.invoiceePrintYN) + vbCrLf
    tmp = tmp + "closeDownState (공급받는자 휴폐업상태) : " + CStr(tiInfo.closeDownState) + vbCrLf
    tmp = tmp + "closeDownStateDate (공급받는자 휴폐업일자 : " + tiInfo.closeDownStateDate + vbCrLf
    tmp = tmp + "trusteeCorpName (수탁자 상호) : " + tiInfo.trusteeCorpName + vbCrLf
    tmp = tmp + "trusteeCorpNum (수탁자 사업자번호) : " + tiInfo.trusteeCorpNum + vbCrLf
    tmp = tmp + "trusteeMgtKey (수탁자 문서관리번호) : " + tiInfo.trusteeMgtKey + vbCrLf
    tmp = tmp + "trusteePrintYN (수탁자 인쇄여부) : " + CStr(tiInfo.trusteePrintYN) + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 대량의 세금계산서 상태/요약 정보를 확인합니다. (최대 1000건)
' - 세금계산서 상태정보(GetInfos API) 응답항목에 대한 자세한 정보는
'  "[전자세금계산서 API 연동매뉴얼] > 4.2. (세금)계산서 상태정보 구성"
'  을 참조하시기 바랍니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '세금계산서 문서관리번호 배열, 최대 1000건
    KeyList.Add "20190207-01"
    KeyList.Add "20190207-02"
    KeyList.Add "20190207-03"
    KeyList.Add "20190207-04"
    
    Set resultList = TaxinvoiceService.GetInfos(txtCorpNum.Text, KeyType, KeyList)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "itemKey(팝빌 관리번호) | taxType (과세형태) | writeDate (작성일자) | regDT (임시저장 일시) | issueType (발행형태) | supplyCostTotal (공급가액 합계) | " + vbCrLf
    tmp = tmp + "taxTotal (세액 합계) | purposeType (영수/청구) |issueDT (발행일시) | lateIssueYN (지연발행 여부) | openYN (개봉 여부) | openDT (개봉 일시) | " + vbCrLf
    tmp = tmp + "stateMemo (상태메모) | stateCode (상태코드) | ntsconfirmNum (국세청승인번호) | ntsresult (국세청 전송결과) | ntssendDT (국세청 전송일시) | " + vbCrLf
    tmp = tmp + "ntsresultDT (국세청 결과 수신일시) | ntssendErrCode (실패사유 사유코드) | modifyCode (수정 사유코드) | interOPYN (연동문서 여부) | invoicerCorpName (공급자 상호) | " + vbCrLf
    tmp = tmp + "invoicerCorpNum (공급자 사업자번호) | invoicerMgtKey (공급자 문서관리번호) | invoicerPrintYN (공급자 인쇄여부) | invoiceeCorpName (공급받는자 상호) | " + vbCrLf
    tmp = tmp + "invoiceeCorpNum (공급받는자 사업자번호) | invoiceeMgtKey(공급받는자 문서관리번호) | invoiceePrintYN(공급받는자 인쇄여부) | closeDownState(공급받는자 휴폐업상태) | " + vbCrLf
    tmp = tmp + "closeDownStateDate(공급받는자 휴폐업일자) | trusteeCorpName (수탁자 상호) | trusteeCorpNum (수탁자 사업자번호) | trusteeMgtKey(수탁자 문서관리번호) | " + vbCrLf
    tmp = tmp + "trusteePrintYN(수탁자 인쇄여부) " + vbCrLf
    
    
    For Each info In resultList
        tmp = tmp + info.itemKey + " | " + info.taxType + " | " + info.writeDate + " | " + info.regDT + " | " + info.issueType + " | " + vbCrLf
        tmp = tmp + info.supplyCostTotal + " | " + info.taxTotal + " | " + info.purposeType + " | " + info.issueDT + " | " + vbCrLf
        tmp = tmp + info.stateDT + " | " + CStr(info.lateIssueYN) + " | " + CStr(info.openYN) + " | " + info.openDT + " | " + vbCrLf
        tmp = tmp + CStr(info.stateCode) + " | " + info.stateMemo + " | " + info.ntsresult + " | " + info.ntsconfirmNum + " | " + vbCrLf
        tmp = tmp + info.ntssendDT + " | " + info.ntsresultDT + " | " + info.ntssendErrCode + " | " + info.modifyCode + " | " + CStr(info.interOPYN) + " | " + vbCrLf
        tmp = tmp + info.invoicerCorpName + " | " + info.invoicerCorpNum + " | " + info.invoicerMgtKey + " | " + CStr(info.invoicerPrintYN) + " | " + vbCrLf
        tmp = tmp + info.invoiceeCorpName + " | " + info.invoiceeCorpNum + " | " + info.invoiceeMgtKey + " | " + vbCrLf
        tmp = tmp + CStr(info.invoiceePrintYN) + " | " + CStr(info.closeDownState) + " | " + info.closeDownStateDate + " | " + vbCrLf
        tmp = tmp + info.trusteeCorpName + " | " + info.trusteeCorpNum + " | " + info.trusteeMgtKey + " | " + CStr(info.trusteePrintYN) + " | " + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 1건의 세금계산서 상세항목을 확인합니다.
' - 응답항목에 대한 자세한 사항은 "[전자세금계산서 API 연동매뉴얼]
'   > 4.1 (세금)계산서 구성" 을 참조하시기 바랍니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    Set tiDetailInfo = TaxinvoiceService.GetDetailInfo(txtCorpNum.Text, KeyType, txtMgtKey.Text)
     
    If tiDetailInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ntsconfirmNum (국세청 승인번호) : " + tiDetailInfo.ntsconfirmNum + vbCrLf
    tmp = tmp + "issueType (발행형태) : " + tiDetailInfo.issueType + vbCrLf
    tmp = tmp + "taxType (과세형태) : " + tiDetailInfo.taxType + vbCrLf
    tmp = tmp + "chargeDirection (과금방향) : " + tiDetailInfo.chargeDirection + vbCrLf
    tmp = tmp + "serialNum (일련번호) : " + tiDetailInfo.serialNum + vbCrLf
    tmp = tmp + "kwon (권) : " + tiDetailInfo.kwon + vbCrLf
    tmp = tmp + "ho (호) : " + tiDetailInfo.ho + vbCrLf
    tmp = tmp + "writeDate (작성일자) : " + tiDetailInfo.writeDate + vbCrLf
    tmp = tmp + "purposeType (영수/청구) : " + tiDetailInfo.purposeType + vbCrLf
    tmp = tmp + "supplyCostTotal (공급가액 합계) : " + tiDetailInfo.supplyCostTotal + vbCrLf
    tmp = tmp + "taxtotal (세액 합계) : " + tiDetailInfo.taxTotal + vbCrLf
    tmp = tmp + "totalAmount (합계 금액) : " + tiDetailInfo.totalAmount + vbCrLf
    tmp = tmp + "cash (현금) : " + tiDetailInfo.cash + vbCrLf
    tmp = tmp + "chkbill (수표) : " + tiDetailInfo.chkBill + vbCrLf
    tmp = tmp + "credit (외상) : " + tiDetailInfo.credit + vbCrLf
    tmp = tmp + "note (어음) : " + tiDetailInfo.note + vbCrLf
    tmp = tmp + "remark1 (비고1) : " + tiDetailInfo.remark1 + vbCrLf
    tmp = tmp + "remark2 (비고2) : " + tiDetailInfo.remark2 + vbCrLf
    tmp = tmp + "remark3 (비고3) : " + tiDetailInfo.remark3 + vbCrLf
        
    tmp = tmp + "invoicerCorpNum (공급자 사업자번호) : " + tiDetailInfo.invoicerCorpNum + vbCrLf
    tmp = tmp + "invoicerMgtKey (공급자 문서관리번호) : " + tiDetailInfo.invoicerMgtKey + vbCrLf
    tmp = tmp + "invoicerTaxRegID (공급자 종사업장 식별번호) : " + tiDetailInfo.invoicerTaxRegID + vbCrLf
    tmp = tmp + "invoicerCorpName (공급자 상호) : " + tiDetailInfo.invoicerCorpName + vbCrLf
    tmp = tmp + "invoicerCEOName (공급자 대표자 성명) : " + tiDetailInfo.invoicerCEOName + vbCrLf
    tmp = tmp + "invoicerAddr (공급자 주소) : " + tiDetailInfo.invoicerAddr + vbCrLf
    tmp = tmp + "invoicerBizClass (공급자 종목) : " + tiDetailInfo.invoicerBizClass + vbCrLf
    tmp = tmp + "invoicerBizType (공급자 업태) : " + tiDetailInfo.invoicerBizType + vbCrLf
    tmp = tmp + "invoicerContactName (공급자 담당자명) : " + tiDetailInfo.invoicerContactName + vbCrLf
    tmp = tmp + "invoicerTEL (공급자 담당자 연락처) : " + tiDetailInfo.invoicerTEL + vbCrLf
    tmp = tmp + "invoicerHP (공급자 담당자 휴대폰번호) : " + tiDetailInfo.invoicerHP + vbCrLf
    tmp = tmp + "invoicerEmail (공급자 담당자 메일) : " + tiDetailInfo.invoicerEmail + vbCrLf
    tmp = tmp + "invoicerSMSSendYN (발행안내메일 전송여부) : " + CStr(tiDetailInfo.invoicerSMSSendYN) + vbCrLf + vbCrLf
    
    tmp = tmp + "invoiceeCorpNum (공급받는자 사업자번호) : " + tiDetailInfo.invoiceeCorpNum + vbCrLf
    tmp = tmp + "invoiceeType (공급받는자 구분) : " + tiDetailInfo.invoiceeType + vbCrLf
    tmp = tmp + "invoiceeMgtKey (공급받는자 문서관리번호) : " + tiDetailInfo.invoiceeMgtKey + vbCrLf
    tmp = tmp + "invoiceeTaxRegID (공급받는자 종사업장 식별번호) : " + tiDetailInfo.invoiceeTaxRegID + vbCrLf
    tmp = tmp + "invoiceeCorpName (공급받는자 상호) : " + tiDetailInfo.invoiceeCorpName + vbCrLf
    tmp = tmp + "invoiceeCEOName (공급받는자 대표자 성명) : " + tiDetailInfo.invoiceeCEOName + vbCrLf
    tmp = tmp + "invoiceeAddr (공급받는자 주소) : " + tiDetailInfo.invoiceeAddr + vbCrLf
    tmp = tmp + "invoiceeBizClass (공급받는자 종목) : " + tiDetailInfo.invoiceeBizClass + vbCrLf
    tmp = tmp + "invoiceeBizType (공급받는자 업태) : " + tiDetailInfo.invoiceeBizType + vbCrLf
    tmp = tmp + "invoiceeContactName1 (공급받는자 담당자명) : " + tiDetailInfo.invoiceeContactName1 + vbCrLf
    tmp = tmp + "invoiceeTEL1 (공급받는자 담당자 연락처) : " + tiDetailInfo.invoiceeTEL1 + vbCrLf
    tmp = tmp + "invoiceeHP1 (공급받는자 담당자 휴대폰번호) : " + tiDetailInfo.invoiceeHP1 + vbCrLf
    tmp = tmp + "invoiceeEmail1 (공급받는자 담당자 메일) : " + tiDetailInfo.invoiceeEmail1 + vbCrLf
    tmp = tmp + "closeDownState (공급받는자 휴폐업상태) : " + CStr(tiDetailInfo.closeDownState) + vbCrLf
    tmp = tmp + "closeDownStateDate (공급받는자 휴폐업일자) : " + tiDetailInfo.closeDownStateDate + vbCrLf + vbCrLf

    tmp = tmp + "modifyCode(수정사유 코드) : " + tiDetailInfo.modifyCode + vbCrLf
    tmp = tmp + "orgNTSConfirmNum(원본 세금계산서 국세청승인번호) : " + tiDetailInfo.orgNTSConfirmNum + vbCrLf
    tmp = tmp + "originalTaxinvoiceKey(원본 팝빌 관리번호) : " + tiDetailInfo.originalTaxinvoiceKey + vbCrLf
   
    If (tiDetailInfo.detailList Is Nothing) = False Then
        For Each detail In tiDetailInfo.detailList
            tmp = tmp + "serialNum (일련번호) : " + CStr(detail.serialNum) + vbCrLf
            tmp = tmp + "purchaseDT (거래일자) : " + detail.purchaseDT + vbCrLf
            tmp = tmp + "itemName (품명) : " + detail.itemName + vbCrLf
            tmp = tmp + "spec (규격) : " + detail.spec + vbCrLf
            tmp = tmp + "qty (수량) : " + detail.qty + vbCrLf
            tmp = tmp + "unitcost (단가) : " + detail.unitCost + vbCrLf
            tmp = tmp + "supplycost (공급가액) : " + detail.supplyCost + vbCrLf
            tmp = tmp + "tax (세액) : " + detail.tax + vbCrLf
            tmp = tmp + "remark (비고) : " + detail.remark + vbCrLf + vbCrLf
        Next
    End If
    
    If (tiDetailInfo.addContactList Is Nothing) = False Then
        For Each contact In tiDetailInfo.addContactList
            tmp = tmp + "serialNum (일련번호) : " + CStr(contact.serialNum) + vbCrLf
            tmp = tmp + "contactName (담당자 성명) : " + contact.ContactName + vbCrLf
            tmp = tmp + "email (이메일주소) : " + contact.email + vbCrLf + vbCrLf
        Next
    End If
    
    MsgBox tmp
End Sub

'=========================================================================
' 검색조건을 사용하여 세금계산서 목록을 조회합니다.
' - 응답항목에 대한 자세한 사항은 "[전자세금계산서 API 연동매뉴얼] >
'   3.5.4. Search(목록 조회)"를 참조하시기 바랍니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    
    '[필수] 일자유형, R-등록일시 W-작성일자 I-발행일시 중 택1
    DType = "W"
    
    '[필수] 시작일자, yyyyMMdd
    SDate = "20190101"
    
    '[필수] 종료일자, yyyyMMdd
    EDate = "20190201"
    
    '전송상태값 배열, 미기재시 전체상태조회, 문서상태값 3자리숫자 작성 2,3번째 와일드카드 가능
    '상태코드에 대한 자세한 사항은 "[전자세금계산서 API 연동매뉴얼] > 5.1 세금계산서 상태코드" 를 참조하시기 바랍니다.
    state.Add "3**"
    state.Add "6**"
    
    '문서유형 배열, N-일반 M-수정 중 선택, 미기재시 전체조회
    TType.Add "N"
    TType.Add "M"
    
    '과세형태 배열, T-과세, N-면세 Z-영세 중 선택, 미기재시 전체조회
    taxType.Add "T"
    taxType.Add "N"
    taxType.Add "Z"
    
    '발행형태 배열, N-정발행, R-역발행 T-위수탁
    issueType.Add "N"
    issueType.Add "R"
    issueType.Add "T"
    
    '지연발행 여부, 0-정상발행 조회 1-지연발행 조회, 공백처리시 전체조회
    LateOnly = ""
    
    '페이지번호, 기본값 ‘1
    Page = 1
    
    '페이지당 검색개수, 기본값 500, 최대 1000
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
    
    '연동문서 조회 여부, 공백-전체조회, 0-일반문서 조회, 1-연동문서 조회
    interOPYN = ""
    
    Set tiSearchList = TaxinvoiceService.Search(txtCorpNum.Text, KeyType, DType, SDate, EDate, state, _
                    TType, taxType, LateOnly, Page, PerPage, Order, TaxRegIDType, TaxRegID, TaxRegIDYN, QString, _
                    txtUserID.Text, interOPYN, issueType)
     
    If tiSearchList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "code (응답코드) : " + CStr(tiSearchList.code) + vbCrLf
    tmp = tmp + "total (총 검색결과 건수) : " + CStr(tiSearchList.total) + vbCrLf
    tmp = tmp + "perPage (페이지당 검색개수) : " + CStr(tiSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum (페이지 번호) : " + CStr(tiSearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (페이지 개수) : " + CStr(tiSearchList.pageCount) + vbCrLf
    tmp = tmp + "message (응답메시지) : " + tiSearchList.message + vbCrLf + vbCrLf
    
    tmp = tmp + "itemKey(팝빌 관리번호) |  taxType (과세형태) |  writeDate (작성일자) |  regDT (임시저장 일시) |  issueType (발행형태) |  supplyCostTotal (공급가액 합계) | " + _
         "taxTotal (세액 합계) |  purposeType (영수/청구) | issueDT (발행일시) | lateIssueYN (지연발행 여부) | openYN (개봉 여부) | openDT (개봉 일시) | " + _
         "stateMemo (상태메모) | stateCode (상태코드) | ntsconfirmNum (국세청승인번호) | ntsresult (국세청 전송결과) | ntssendDT (국세청 전송일시) | " + _
         "ntsresultDT (국세청 결과 수신일시) | ntssendErrCode (전송실패 사유코드) | modifyCode (수정 사유코드) | interOPYN (연동문서 여부) | invoicerCorpName (공급자 상호) | " + _
         "invoicerCorpNum (공급자 사업자번호) | invoicerMgtKey (공급자 문서관리번호) | invoicerPrintYN (공급자 인쇄여부) | invoiceeCorpName (공급받는자 상호) | " + _
         "invoiceeCorpNum (공급받는자 사업자번호) | invoiceeMgtKey(공급받는자 문서관리번호) | invoiceePrintYN(공급받는자 인쇄여부) | closeDownState(공급받는자 휴폐업상태) |" + _
         "closeDownStateDate(공급받는자 휴폐업일자) | trusteeCorpName (수탁자 상호) | trusteeCorpNum (수탁자 사업자번호) | trusteeMgtKey(수탁자 문서관리번호) | " + _
         "trusteePrintYN(수탁자 인쇄여부) " + vbCrLf + vbCrLf
            
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
        tmp = tmp + info.ntsconfirmNum + " | "
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
' 세금계산서 상태 변경이력을 확인합니다.
' - 상태 변경이력 확인(GetLogs API) 응답항목에 대한 자세한 정보는
'   "[전자세금계산서 API 연동매뉴얼] > 3.5.5 상태 변경이력 확인"
'   을 참조하시기 바랍니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    
    Set resultList = TaxinvoiceService.GetLogs(txtCorpNum.Text, KeyType, txtMgtKey.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "DocLogType(로그타입) | Log(이력정보) | ProcType(처리형태) | procCorpName(처리회사명) | procContactName(처리담당자) | " _
        + "ProcMemo(처리메모) | RegDT(등록일시) | IP(아이피) " + vbCrLf
    
    For Each log In resultList
        tmp = tmp + CStr(log.docLogType) + " | " + log.log + " | " + log.procType + " | " + log.procCorpName + " | " + log.procContactName _
        + log.procMemo + " | " + log.regDT + " | " + log.ip + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 발행 안내메일을 재전송합니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '수신메일주소
    receiverEmail = "test@test.com"
    
    Set Response = TaxinvoiceService.SendEmail(txtCorpNum.Text, KeyType, txtMgtKey.Text, receiverEmail)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 알림문자를 전송합니다. (단문/SMS- 한글 최대 45자)
' - 알림문자 전송시 포인트가 차감됩니다. (전송실패시 환불처리)
' - 전송내역 확인은 "팝빌 로그인" > [문자 팩스] > [문자] > [전송내역] 탭에서
'   전송결과를 확인할 수 있습니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    ' 발신번호
    sendNum = "07043042991"
    
    ' 수신번호
    receiveNum = "010-111-222"
    
    ' 메시지 내용, 최대 90Byte (한글 45자), 길이를 초과한 내용은 삭제되어 전송됩니다.
    Contents = "링크허브에서 세금계산서를 발행하였습니다. 메일확인 바랍니다."
        
    
    Set Response = TaxinvoiceService.SendSMS(txtCorpNum.Text, KeyType, txtMgtKey.Text, _
                            sendNum, receiveNum, Contents)
    
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '발신자 번호
    sendNum = "07043042991"
    
    '수신자 팩스 번호
    receiveNum = "010-222-4444"
    
    Set Response = TaxinvoiceService.SendFax(txtCorpNum.Text, KeyType, txtMgtKey.Text, sendNum, receiveNum)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 팝빌 사이트에서 작성한 세금계산서에 파트너 문서관리번호를 할당합니다.
' - 문서관리번호가 존재하지 않는 세금계산서만 할당이 가능 합니다.
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    '세금계산서 아이템키, 목록조회(Search) API의 반환항목중 ItemKey 참조
    itemKey = "018090515070600001"
            
    '할당할 문서관리번호, 숫자, 영문, '-', '_' 조합으로
    '1~24자리까지 사업자번호별 중복없는 고유번호 할당
    MgtKey = "20190201-001"
        
    Set Response = TaxinvoiceService.AssignMgtKey(txtCorpNum.Text, KeyType, itemKey, MgtKey)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
'전자세금계산서에 1건의 전자명세서를 첨부합니다.
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
    SubMgtKey = "20190207-01"
        
    Set Response = TaxinvoiceService.AttachStatement(txtCorpNum.Text, KeyType, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
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
    
    '첨부해제할 전자명세서 종류코드, 121-거래명세서, 122-청구서, 123-견적서, 124-발주서, 125-입금표, 126-영수증
    SubItemCode = 121
    
    '첨부해제할 전자명세서 관리번호
    SubMgtKey = "20190207-01"

    Set Response = TaxinvoiceService.DetachStatement(txtCorpNum.Text, KeyType, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 전자세금계산서 관련 메일전송 항목에 대한 전송여부를 목록으로 반환합니다
'=========================================================================
Private Sub btnListemailconfig_Click(index As Integer)
    Dim resultList As Collection
    Dim i As Integer
    
    Set resultList = TaxinvoiceService.ListEmailConfig(txtCorpNum.Text, txtUserID.Text)
    
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
 
    Dim tmp As String
    
    tmp = "메일전송유형(EmailType) | 전송여부(SendYN) " + vbCrLf
    
    Dim info As PBEmailConfig
    
    For i = 1 To resultList.Count
        If resultList(i).emailType = "TAX_ISSUE" Then
            tmp = tmp + "[정발행] 공급받는자에게 전자세금계산서 발행 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_ISSUE_INVOICER" Then
            tmp = tmp + "[정발행] 공급자에게 전자세금계산서 발행 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_CHECK" Then
            tmp = tmp + "[정발행] 공급자에게 전자세금계산서 수신확인 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_CANCEL_ISSUE" Then
            tmp = tmp + "[정발행] 공급받는자에게 전자세금계산서 발행취소 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
                    
        If resultList(i).emailType = "TAX_SEND" Then
            tmp = tmp + "[발행예정] 공급받는자에게 [발행예정] 세금계산서 발송 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_ACCEPT" Then
            tmp = tmp + "[발행예정] 공급자에게 [발행예정] 세금계산서 승인 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_ACCEPT_ISSUE" Then
            tmp = tmp + "[발행예정] 공급자에게 [발행예정] 세금계산서 자동발행 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_DENY" Then
            tmp = tmp + "[발행예정] 공급자에게 [발행예정] 세금계산서 거부 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If

        If resultList(i).emailType = "TAX_CANCEL_SEND" Then
            tmp = tmp + "[발행예정] 공급받는자에게 [발행예정] 세금계산서 취소 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
                    
        If resultList(i).emailType = "TAX_REQUEST" Then
            tmp = tmp + "[역발행] 공급자에게 세금계산서를 발행요청 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_CANCEL_REQUEST" Then
            tmp = tmp + "[역발행] 공급받는자에게 세금계산서 취소 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
            If resultList(i).emailType = "TAX_REFUSE" Then
            tmp = tmp + "[역발행] 공급받는자에게 세금계산서 거부 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_TRUST_ISSUE" Then
            tmp = tmp + "[위수탁발행] 공급받는자에게 전자세금계산서 발행 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_TRUST_ISSUE_TRUSTEE" Then
            tmp = tmp + "[위수탁발행] 수탁자에게 전자세금계산서 발행 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
                    
        If resultList(i).emailType = "TAX_TRUST_ISSUE_INVOICER" Then
            tmp = tmp + "[위수탁발행] 공급자에게 전자세금계산서 발행 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_TRUST_CANCEL_ISSUE" Then
            tmp = tmp + "[위수탁발행] 공급받는자에게 전자세금계산서 발행취소 알림 : " + resultList(i).emailType + " | "
          tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
    
        If resultList(i).emailType = "TAX_TRUST_SEND" Then
            tmp = tmp + "[위수탁발행] 공급자에게 전자세금계산서 발행취소 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_TRUST_CANCEL_ISSUE_INVOICER" Then
            tmp = tmp + "[위수탁발행예정] 공급받는자에게 [발행예정] 세금계산서 발송 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_TRUST_ACCEPT" Then
            tmp = tmp + "[위수탁발행예정] 수탁자에게 [발행예정] 세금계산서 승인 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
                    
        If resultList(i).emailType = "TAX_TRUST_ACCEPT_ISSUE" Then
            tmp = tmp + "[위수탁발행예정] 수탁자에게 [발행예정] 세금계산서 자동발행 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_TRUST_DENY" Then
            tmp = tmp + "[위수탁발행예정] 수탁자에게 [발행예정] 세금계산서 거부 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_TRUST_CANCEL_SEND" Then
            tmp = tmp + "[위수탁발행예정] 공급받는자에게 [발행예정] 세금계산서 취소 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_CLOSEDOWN" Then
            tmp = tmp + "[처리결과] 거래처의 휴폐업 여부 확인 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "TAX_NTSFAIL_INVOICER" Then
            tmp = tmp + "[처리결과] 전자세금계산서 국세청 전송실패 안내) : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
                    
        If resultList(i).emailType = "TAX_SEND_INFO" Then
            tmp = tmp + "[정기발송] 전월 귀속분 [매출 발행 대기] 세금계산서 발행 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "ETC_CERT_EXPIRATION" Then
            tmp = tmp + "[정기발송] 팝빌에서 이용중인 공인인증서의 갱신 알림 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 전자세금계산서 관련 메일전송 항목에 대한 전송여부를 수정합니다.
'
' 메일전송유형
' [정발행]
' TAX_ISSUE : 공급받는자에게 전자세금계산서가 발행 되었음을 알려주는 메일입니다.
' TAX_ISSUE_INVOICER : 공급자에게 전자세금계산서가 발행 되었음을 알려주는 메일입니다.
' TAX_CHECK : 공급자에게 전자세금계산서가 수신확인 되었음을 알려주는 메일입니다.
' TAX_CANCEL_ISSUE : 공급받는자에게 전자세금계산서가 발행취소 되었음을 알려주는 메일입니다.
'
' [발행예정]
' TAX_SEND : 공급받는자에게 [발행예정] 세금계산서가 발송 되었음을 알려주는 메일입니다.
' TAX_ACCEPT : 공급자에게 [발행예정] 세금계산서가 승인 되었음을 알려주는 메일입니다.
' TAX_ACCEPT_ISSUE : 공급자에게 [발행예정] 세금계산서가 자동발행 되었음을 알려주는 메일입니다.
' TAX_DENY : 공급자에게 [발행예정] 세금계산서가 거부 되었음을 알려주는 메일입니다.
' TAX_CANCEL_SEND : 공급받는자에게 [발행예정] 세금계산서가 취소 되었음을 알려주는 메일입니다.
'
' [역발행]
' TAX_REQUEST : 공급자에게 세금계산서를 전자서명 하여 발행을 요청하는 메일입니다.
' TAX_CANCEL_REQUEST : 공급받는자에게 세금계산서가 취소 되었음을 알려주는 메일입니다.
' TAX_REFUSE : 공급받는자에게 세금계산서가 거부 되었음을 알려주는 메일입니다.
'
' [위수탁발행]
' TAX_TRUST_ISSUE : 공급받는자에게 전자세금계산서가 발행 되었음을 알려주는 메일입니다.
' TAX_TRUST_ISSUE_TRUSTEE : 수탁자에게 전자세금계산서가 발행 되었음을 알려주는 메일입니다.
' TAX_TRUST_ISSUE_INVOICER : 공급자에게 전자세금계산서가 발행 되었음을 알려주는 메일입니다.
' TAX_TRUST_CANCEL_ISSUE : 공급받는자에게 전자세금계산서가 발행취소 되었음을 알려주는 메일입니다.
' TAX_TRUST_CANCEL_ISSUE_INVOICER : 공급자에게 전자세금계산서가 발행취소 되었음을 알려주는 메일입니다.
'
' [위수탁 발행예정]
' TAX_TRUST_SEND : 공급받는자에게 [발행예정] 세금계산서가 발송 되었음을 알려주는 메일입니다.
' TAX_TRUST_ACCEPT : 수탁자에게 [발행예정] 세금계산서가 승인 되었음을 알려주는 메일입니다.
' TAX_TRUST_ACCEPT_ISSUE : 수탁자에게 [발행예정] 세금계산서가 자동발행 되었음을 알려주는 메일입니다.
' TAX_TRUST_DENY : 수탁자에게 [발행예정] 세금계산서가 거부 되었음을 알려주는 메일입니다.
' TAX_TRUST_CANCEL_SEND : 공급받는자에게 [발행예정] 세금계산서가 취소 되었음을 알려주는 메일입니다.
'
' [처리결과]
' TAX_CLOSEDOWN : 거래처의 휴폐업 여부를 확인하여 안내하는 메일입니다.
' TAX_NTSFAIL_INVOICER : 전자세금계산서 국세청 전송실패를 안내하는 메일입니다.
'
' [정기발송]
' TAX_SEND_INFO : 전월 귀속분 [매출 발행 대기] 세금계산서의 발행을 안내하는 메일입니다.
' ETC_CERT_EXPIRATION : 팝빌에서 이용중인 공인인증서의 갱신을 안내하는 메일입니다.
'
'=========================================================================
Private Sub btnUpdateemailconfig_Click(index As Integer)
    Dim Response As PBResponse
    Dim emailType As String
    Dim sendYN As Boolean
    
    '메일 전송 유형
    emailType = "TAX_ISSUE"

    '전송 여부 (True = 전송, False = 미전송)
    sendYN = True
    
    Set Response = TaxinvoiceService.UpdateEmailConfig(txtCorpNum.Text, emailType, sendYN, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
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
    
    url = TaxinvoiceService.GetPopUpURL(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
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
    
    url = TaxinvoiceService.GetPrintURL(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 1건의 전자세금계산서 인쇄(공급받는자) URL을 반환합니다.
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
    
    url = TaxinvoiceService.GetEPrintURL(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 대량의 전자세금계산서 인쇄팝업 URL을 반환합니다. (최대 100건)
' 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
'=========================================================================
Private Sub btnGetMassPrintURL_Click()
    Dim url As String
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
            MsgBox "관리번호 형태를 선택해주세요."
            Exit Sub
    End Select
    
    ' 전자세금계산서 문서 관리번호 배열 (최대 100건)
    KeyList.Add "20190207-01"
    KeyList.Add "20190207-02"
    KeyList.Add "20190207-03"
    KeyList.Add "20190207-04"
    
    url = TaxinvoiceService.GetMassPrintURL(txtCorpNum.Text, KeyType, KeyList)
     
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 공급받는자 수신메일 링크주소를 반환합니다.
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
    
    url = TaxinvoiceService.GetMailURL(txtCorpNum.Text, KeyType, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 팝빌 > 전자세금계산서 > 임시(연동)문서함 팝업 URL을 반환합니다.
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
' 팝빌 > 전자세금계산서 > 매출 문서함 팝업 URL을 반환합니다.
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
' 팝빌 > 전자세금계산서 > 매입 문서함 팝업 URL을 반환합니다.
' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
'=========================================================================
Private Sub btnGetURL_PBOX_Click()
    Dim url As String
    
    url = TaxinvoiceService.GetURL(txtCorpNum.Text, txtUserID.Text, "PBOX")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(TaxinvoiceService.LastErrCode) + vbCrLf + "응답메시지 : " + TaxinvoiceService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 팝빌 > 전자세금계산서 > 매출 문서작성 팝업 URL을 반환합니다.
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

Private Sub Form_Load()
    ' 모듈 초기화
    TaxinvoiceService.Initialize LinkID, SecretKey

    ' 연동환경 설정값 True(개발용), False(상업용), 상업용 전환시 False로 변경.
    TaxinvoiceService.IsTest = True

    cboMgtKeyType.AddItem "SELL"
    cboMgtKeyType.AddItem "BUY"
    cboMgtKeyType.AddItem "TRUSTEE"
End Sub
