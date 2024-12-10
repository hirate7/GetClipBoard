﻿Imports GetClipBoard.CMyNumAPI2

Public Class CJACCESS3

    Public Shared adoMyNumClip As ADODB.Connection
    Public Shared MyNum_Connect As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + CGetClipBoard.Path_KDATA + "MyNumClip.mdb;Jet OLEDB:Database Password=4300365;Persist Security Info=False"

    Public Shared Sub Open_Connection()

        Dim lngOptions As String

        lngOptions = ADODB.ConnectOptionEnum.adAsyncConnect '非同期接続

        adoMyNumClip = New ADODB.Connection

        adoMyNumClip.Mode = ADODB.ConnectModeEnum.adModeReadWrite
        adoMyNumClip.Open(MyNum_Connect, , , lngOptions)

        '/// 接続が完了するまで待機する ///
        Do While adoMyNumClip.State <> ADODB.ObjectStateEnum.adStateOpen
            'System.Windows.Forms.Application.DoEvents()
        Loop

    End Sub

    Public Shared Sub Close_Connection()

        adoMyNumClip.Close()

    End Sub

    Public Shared Sub Add_ClipData(CP As CClipData)

        Dim Rs As New ADODB.Recordset
        Dim SQL As String

        adoMyNumClip.BeginTrans()

        On Error GoTo TrnFalse

        SQL = "select * from ClipData where QualificationConfirmationDate='" + CP.QualificationConfirmationDate + "'" +
            " and InsurerNumber='" + CP.InsurerNumber + "'" +
            " and InsuredCardSymbol='" + CP.InsuredCardSymbol + "'" +
            " and InsuredIdentificationNumber='" + CP.InsuredIdentificationNumber + "'" +
            " and InsuredBranchNumber='" + CP.InsuredBranchNumber + "'"

        Rs.Open(SQL, adoMyNumClip, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

        If Rs.BOF And Rs.EOF Then
            Rs.AddNew()
        Else
            Rs.MoveFirst()
        End If

        With CP

            '処理実行日時				
            Rs.Fields("ProcessExecutionTime").Value = .ProcessExecutionTime
            '資格確認日				
            Rs.Fields("QualificationConfirmationDate").Value = .QualificationConfirmationDate
            '資格有効性				
            Rs.Fields("QualificationValidity").Value = .QualificationValidity

            '被保険者証区分			
            Rs.Fields("InsuredCardClassification").Value = .InsuredCardClassification
            '保険者番号
            Rs.Fields("InsurerNumber").Value = .InsurerNumber
            '被保険者証記号			
            Rs.Fields("InsuredCardSymbol").Value = .InsuredCardSymbol
            '被保険者証番号
            Rs.Fields("InsuredIdentificationNumber").Value = .InsuredIdentificationNumber
            '被保険者証枝番			
            Rs.Fields("InsuredBranchNumber").Value = .InsuredBranchNumber
            '本人・家族の別			
            Rs.Fields("PersonalFamilyClassification").Value = .PersonalFamilyClassification
            '被保険者氏名(世帯主氏名)			
            Rs.Fields("InsuredName").Value = .InsuredName
            '氏名			
            Rs.Fields("Name").Value = .Name
            '氏名（その他）			
            Rs.Fields("NameOfOther").Value = .NameOfOther
            '氏名カナ			
            Rs.Fields("NameKana").Value = .NameKana
            '氏名カナ（その他）			
            Rs.Fields("NameOfOtherKana").Value = .NameOfOtherKana
            '性別1			
            Rs.Fields("Sex1").Value = .Sex1
            '性別2		
            Rs.Fields("Sex2").Value = .Sex2
            '生年月日
            Rs.Fields("Birthdate").Value = .Birthdate
            '住所		
            Rs.Fields("Address").Value = .Address
            '郵便番号
            Rs.Fields("PostNumber").Value = .PostNumber
            '資格取得年月日			
            Rs.Fields("QualificationDate").Value = .QualificationDate
            '被保険者証有効開始年月日			
            Rs.Fields("InsuredCardValidDate").Value = .InsuredCardValidDate
            '被保険者証有効終了年月日		
            Rs.Fields("InsuredCardExpirationDate").Value = .InsuredCardExpirationDate
            '被保険者証一部負担金割合			
            Rs.Fields("InsuredPartialContributionRatio").Value = .InsuredPartialContributionRatio
            '未就学区分			
            Rs.Fields("PreschoolClassification").Value = .PreschoolClassification
            '保険者名称			
            Rs.Fields("InsurerName").Value = .InsurerName

            '高齢受給者証有効開始年月日
            Rs.Fields("ElderlyRecipientValidStartDate").Value = .ElderlyRecipientValidStartDate
            '高齢受給者証有効終了年月日		
            Rs.Fields("ElderlyRecipientValidEndDate").Value = .ElderlyRecipientValidEndDate
            '高齢受給者証一部負担金割合		
            Rs.Fields("ElderlyRecipientContributionRatio").Value = .ElderlyRecipientContributionRatio

            Rs.Fields("RawData").Value = .RawData

        End With

        Rs.Update()

        Rs.Close()

        adoMyNumClip.CommitTrans()

        Beep()

        Exit Sub

trnfalse:

        adoMyNumClip.RollbackTrans()

    End Sub


End Class