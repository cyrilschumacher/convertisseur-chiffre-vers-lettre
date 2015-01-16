Function ChiffreEnLettre(chiffre As Double) As String
    ChiffreSansVirgule = Val(chiffre)
    If ChiffreSansVirgule <> 0 Then
        LongueurChiffre = Len(ChiffreSansVirgule)
        If LongueurChiffre >= 4 And LongueurChiffre <= 6 Then
            EnTouteLettre = ConvertirMillier(chiffre / 1000)
            EnTouteLettre = EnTouteLettre & Convertir(chiffre Mod 1000)
        ElseIf LongueurChiffre = 3 Then
            EnTouteLettre = ConvertirCentaine(ChiffreSansVirgule)
        ElseIf LongueurChiffre = 2 Then
            EnTouteLettre = ConvertirDizaine(ChiffreSansVirgule)
        Else
            EnTouteLettre = ConvertirUnite(ChiffreSansVirgule)
        End If
        
        EnTouteLettre = Trim(EnTouteLettre)
    End If
    
    Virgule = ApresVirgule(chiffre)
    If Virgule <> 0 Then
        Virgule = Convertir(CStr(Virgule))
        EnTouteLettre = Trim(EnTouteLettre) & " " & Virgule
        If Virgule = "un" Then
            EnTouteLettre = EnTouteLettre & " CENTIME"
        Else
            EnTouteLettre = EnTouteLettre & " CENTIMES"
        End If
    End If
    
    ChiffreEnLettre = UCase(EnTouteLettre)
End Function

Private Function ApresVirgule(chiffre As Double) As String
    chiffre = Format(chiffre, "###0.00")
    ChiffreSansVirgule = Val(chiffre)
    
    Virgule = Format(chiffre - ChiffreSansVirgule, "###0.00")
    LongueurVirgule = Len(Virgule) - 2
    AVirgule = Virgule * (10 ^ LongueurVirgule)
    ApresVirgule = AVirgule
End Function

Private Function Convertir(chiffre As String) As String
    Longueur = Len(chiffre)
    If Longueur = 1 Then
        Convertir = ConvertirUnite(chiffre)
    ElseIf Longueur = 2 Then
        Convertir = ConvertirDizaine(chiffre)
    ElseIf Longueur = 3 Then
        Convertir = ConvertirCentaine(chiffre)
    End If
End Function

Private Function ConvertirUnite(chiffre) As String
    Unite = Array("", "un", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf")
    ConvertirUnite = Unite(chiffre)
End Function

Private Function ConvertirDizaine(chiffre) As String
    Dizaine = Array("dix", "onze", "douze", "treize", "quartoze", "quinze", "seize", "dix-sept", "dix-huit", "dix-neuf")
    Dizaine_2 = Array("", "dix", "vingt", "trente", "quarante", "cinquante", "soixante", "soixante", "quatre-vingt", "quatre-vingt")
    
    PremierNombre = Left(chiffre, 1)
    DeuxiemeNombre = Right(chiffre, 1)
    
    If PremierNombre = 1 Then
        EnTouteLettre = "" & Dizaine(PremierNombre)
    Else
        EnTouteLettre = Dizaine_2(PremierNombre)
        If PremierNombre = 9 Or PremierNombre = 7 Then
            EnTouteLettre = " " & EnTouteLettre & " " & Dizaine(DeuxiemeNombre)
        Else
            EnTouteLettre = " " & EnTouteLettre & " " & ConvertirUnite(DeuxiemeNombre)
        End If
    End If
    
    ConvertirDizaine = EnTouteLettre
End Function

Private Function ConvertirCentaine(chiffre) As String
    PremierNombre = Left(chiffre, 1)
    If PremierNombre <> 0 Then
        If PremierNombre = 1 Then
            EnTouteLettre = " cent"
        Else
            EnTouteLettre = " " & ConvertirUnite(PremierNombre) & " cents"
        End If
        
        DeuxiemeNombre = Right(chiffre, 2)
        EnTouteLettre = EnTouteLettre & ConvertirDizaine(DeuxiemeNombre)
        ConvertirCentaine = EnTouteLettre
    End If
End Function

Private Function ConvertirMillier(chiffre) As String
    chiffre = Val(chiffre)
    LongueurChiffre = Len(chiffre)
    PremierNombre = Left(chiffre, LongueurChiffre)
    If PremierNombre = 1 And LongueurChiffre = 1 Then
        ConvertirMillier = "mille"
    Else
        If LongueurChiffre = 3 Then
            EnTouteLettre = ConvertirCentaine(PremierNombre)
        ElseIf LongueurChiffre = 2 Then
            EnTouteLettre = ConvertirDizaine(PremierNombre)
        Else
            EnTouteLettre = ConvertirUnite(PremierNombre)
        End If
        ConvertirMillier = EnTouteLettre & " milles"
    End If
End Function
