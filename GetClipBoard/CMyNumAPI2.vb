Public Class CMyNumAPI2

    Public Class CClipData

        '処理実行日時
        Public ProcessExecutionTime As DateTime

        '資格確認日
        Public QualificationConfirmationDate As String

        '資格有効性
        Public QualificationValidity As String

        '被保険者証区分			
        Public InsuredCardClassification As String
        '保険者番号
        Public InsurerNumber As String = ""
        '被保険者証記号			
        Public InsuredCardSymbol As String = ""
        '被保険者証番号
        Public InsuredIdentificationNumber As String = ""
        '被保険者証枝番			
        Public InsuredBranchNumber As String = ""
        '本人・家族の別			
        Public PersonalFamilyClassification As String
        '被保険者氏名(世帯主氏名)			
        Public InsuredName As String
        '氏名			
        Public Name As String
        '氏名（その他）			
        Public NameOfOther As String
        '氏名カナ			
        Public NameKana As String
        '氏名カナ（その他）			
        Public NameOfOtherKana As String
        '性別1			
        Public Sex1 As String
        '性別2		
        Public Sex2 As String
        '生年月日
        Public Birthdate As String
        '住所		
        Public Address As String
        '郵便番号
        Public PostNumber As String
        '資格取得年月日			
        Public QualificationDate As String
        '被保険者証交付年月日
        Public InsuredCertificateIssuanceDate As String
        '被保険者証有効開始年月日			
        Public InsuredCardValidDate As String
        '被保険者証有効終了年月日		
        Public InsuredCardExpirationDate As String
        '被保険者証一部負担金割合			
        Public InsuredPartialContributionRatio As String
        '未就学区分			
        Public PreschoolClassification As String
        '保険者名称			
        Public InsurerName As String

        '高齢受給者証有効開始年月日
        Public ElderlyRecipientValidStartDate As String
        '高齢受給者証有効終了年月日		
        Public ElderlyRecipientValidEndDate As String
        '高齢受給者証一部負担金割合		
        Public ElderlyRecipientContributionRatio As String

        '照会番号
        Public ReferenceNumber As String

        Public RawData As String

    End Class

End Class
