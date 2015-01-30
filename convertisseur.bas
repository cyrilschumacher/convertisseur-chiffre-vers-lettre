''' The MIT License (MIT)
'''
''' Copyright (c) 2014 Cyril Schumacher.fr
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to deal
''' in the Software without restriction, including without limitation the rights
''' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
''' copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in all
''' copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
''' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
''' SOFTWARE.

Option Explicit

''' <summary>
'''     Convertit un nombre en toutes lettres.
''' </summary>
''' <param name="chiffre">Nombre à convertir.</param>
''' <returns>Une chaîne de caractère représentant le nombre.</returns>
Function ChiffreEnLettre(chiffre As Double) As String
    ''' Obtient le chiffre sans virgule.
    Dim ChiffreSansVirgule As Variant
    ChiffreSansVirgule = Val(chiffre)
    
    If ChiffreSansVirgule <> 0 Then
        ''' Obtient la taille du chiffre.
        Dim LongueurChiffre As Long
        LongueurChiffre = Len(ChiffreSansVirgule)
        
        Dim EnTouteLettre As String
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
    
    Dim Virgule As String
    Virgule = ApresVirgule(chiffre)
    If Virgule <> 0 Then
        Virgule = Convertir(CStr(Virgule))
        EnTouteLettre = Trim(EnTouteLettre) & " " & Virgule
    End If
    
    ChiffreEnLettre = UCase(EnTouteLettre)
End Function

''' <summary>
'''     Convertit le nombre situé après la virgule.
''' </summary>
''' <param name="chiffre">Nombre à convertir.</param>
''' <returns>Une chaîne de caractère représentant le nombre.</returns>
Private Function ApresVirgule(chiffre As Double) As String
    chiffre = Format(chiffre, "###0.00")
    
    Dim ChiffreSansVirgule As Variant
    ChiffreSansVirgule = Val(chiffre)
    
    Dim Virgule As String
    Virgule = Format(chiffre - ChiffreSansVirgule, "###0.00")
    
    Dim LongueurVirgule As Long
    LongueurVirgule = Len(Virgule) - 2
    
    Dim AVirgule As Variant
    AVirgule = Virgule * (10 ^ LongueurVirgule)
    ApresVirgule = AVirgule
End Function

''' <summary>
'''     Convertit un nombre en toutes lettres.
''' </summary>
''' <param name="chiffre">Nombre à convertir.</param>
''' <returns>Une chaîne de caractère représentant le nombre.</returns>
Private Function Convertir(chiffre As String) As String
    Dim Longueur As Long
    Longueur = Len(chiffre)
    
    If Longueur = 1 Then
        Convertir = ConvertirUnite(chiffre)
    ElseIf Longueur = 2 Then
        Convertir = ConvertirDizaine(chiffre)
    ElseIf Longueur = 3 Then
        Convertir = ConvertirCentaine(chiffre)
    End If
End Function

''' <summary>
'''     Convertit un nombre, représentant une unité, toutes lettres.
''' </summary>
''' <param name="chiffre">Nombre à convertir.</param>
''' <returns>Une chaîne de caractère représentant le nombre.</returns>
Private Function ConvertirUnite(chiffre) As String
    Dim Unite As Variant
    Unite = Array("", "un", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf")
    
    ConvertirUnite = Unite(chiffre)
End Function

''' <summary>
'''     Convertit un nombre, représentant une dizaine, toutes lettres.
''' </summary>
''' <param name="chiffre">Nombre à convertir.</param>
''' <returns>Une chaîne de caractère représentant le nombre.</returns>
Private Function ConvertirDizaine(chiffre) As String
    Dim Dizaine As Variant
    Dizaine = Array("dix", "onze", "douze", "treize", "quartoze", "quinze", "seize", "dix-sept", "dix-huit", "dix-neuf")
    Dim Dizaine_2 As Variant
    Dizaine_2 = Array("", "dix", "vingt", "trente", "quarante", "cinquante", "soixante", "soixante", "quatre-vingt", "quatre-vingt")
    
    Dim PremierNombre As String
    PremierNombre = Left(chiffre, 1)
    Dim DeuxiemeNombre As String
    DeuxiemeNombre = Right(chiffre, 1)
    
    Dim EnTouteLettre As String
    If PremierNombre = 1 Then
        EnTouteLettre = " " & Dizaine(DeuxiemeNombre)
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

''' <summary>
'''     Convertit un nombre, représentant une centaine, toutes lettres.
''' </summary>
''' <param name="chiffre">Nombre à convertir.</param>
''' <returns>Une chaîne de caractère représentant le nombre.</returns>
Private Function ConvertirCentaine(chiffre) As String
    Dim PremierNombre As String
    PremierNombre = Left(chiffre, 1)
    
    If PremierNombre <> 0 Then
        Dim EnTouteLettre As String
        If PremierNombre = 1 Then
            EnTouteLettre = " cent"
        Else
            EnTouteLettre = " " & ConvertirUnite(PremierNombre) & " cents"
        End If
        
        Dim DeuxiemeNombre As String
        DeuxiemeNombre = Right(chiffre, 2)
        EnTouteLettre = EnTouteLettre & ConvertirDizaine(DeuxiemeNombre)
        ConvertirCentaine = EnTouteLettre
    End If
End Function

''' <summary>
'''     Convertit un nombre, représentant un millier, toutes lettres.
''' </summary>
''' <param name="chiffre">Nombre à convertir.</param>
''' <returns>Une chaîne de caractère représentant le nombre.</returns>
Private Function ConvertirMillier(chiffre) As String
    chiffre = Val(chiffre)
    
    Dim LongueurChiffre As Long
    LongueurChiffre = Len(chiffre)
    
    Dim PremierNombre As String
    PremierNombre = Left(chiffre, LongueurChiffre)
    
    If PremierNombre = 1 And LongueurChiffre = 1 Then
        ConvertirMillier = " mille"
    Else
        Dim EnTouteLettre As String
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