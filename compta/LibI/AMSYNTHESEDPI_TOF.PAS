{***********UNITE*************************************************
Auteur  ...... : TGA
Créé le ...... : 10/05/2006
Modifié le ... : 29/05/2006
Description .. : Source TOF de la FICHE : AMSYNTHESEDPI ()
Mots clefs ... : TOF;AMSYNTHESEDPI
Suite......... : MBO - 25/08/2006 - correction pour conseil en compilation
Suite......... : TGA - 05/09/2006 - raccouci sur date et suppression
Suite......... : TGA - 25/09/2006 - nouvel affichage du solde
// MVG 11/10/2006 et 16/10/2006 FQ 18951
Suite......... : MBO - 17/10/2006 - correction pour des pb d'arrondi
Suite......... : MVG - 30/11/2006- FQ 19259
*****************************************************************}
Unit AMSYNTHESEDPI_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     PrintDbg,
     db,
     {$IFNDEF DBXPRESS}
        dbtables,
     {$ELSE}
        uDbxDataSet,
     {$ENDIF}
     mul,
     fe_main,
{$else}
     UtileAGL,
     MaineAGL,
     eMul,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     extctrls,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Grids,
     HTB97,
     Vierge,
     windows,
     Dialogs,
     UTOF,
     HXLSPAS;


Type
  TOF_AMSYNTHESEDPI = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

    private
    { Variables }
    FListe : THGrid ;
    Fradio : THRadioGroup;
    Fcombo : THValComboBox;
    FBvalid : TToolbarButton97;
    FBsupp  : TToolbarButton97;
    FBChercher: TToolbarButton97;
    Ftotal : THLabel;
    SD : TSaveDialog;
    FIndex1 : Integer;
    FIndex2 : Integer;
    FdebN1 : String;
    FfinN1 : String;
    FdebN5 : String;
    FfinN5 : String;
    FDateDebut : TdateTime;
    FDateFin : TdateTime;
    FCadragedroite : String;

    procedure OnClickBchercher(Sender: TObject);
    procedure OnClickRadio(Sender: TObject);
    procedure OnClickBSupprimer(Sender: TObject);
    procedure OnClickCombo(Sender: TObject);
    procedure OnGrilleKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure OnClickBImprimer(Sender: TObject);
    procedure OnClickBExport(Sender: TObject);
    procedure AffichageDeLaListe ;
    procedure AffichageDuSolde ;
    procedure PreparePrintGrid;

  end ;

Function  AMLanceFiche_SYNTHESEDPI(stCode: string):String;
Procedure AM_MAJ_IMMOMVTD(action: String ;nimmo1:String ;nimmo2:String ;Fraction:double);

  // Exemple de lancement
  // Stcode :='01/01/2006;31/12/2006;01/01/2005;31/12/2005;01/01/2004;31/12/2004;01/01/2003;31/12/2003;01/01/2002;31/12/2002;01/01/2001;31/12/2001;01/01/2006;31/12/2006';
  // result:=AGLLanceFiche('AM', 'AMSYNTHESEDPI','' , '', Stcode);


Implementation

const

  HM: array[0..28] of string =
   ('Affectation aux immobilisations',
    'Affectation aux stocks',
    'Affectation aux parts de coopératives agricoles',
    'Réintégration anticipée',
    'Réintégration obligatoire',
    'Affectation à l''immobilisation',
    'Affectation aux immobilisations sorties',
    'Tous les exercices',
    'De ',
    ' au ',
    'DPI groupées par motif',
    'DPI détaillées par motif',
    'DPI par date d''investissement',
    '                                                             Total ',
    ' (investissement du ',
    ' du ',
    'Confirmez-vous la suppression de l''affectation des DPI faite sur l''immo n°',
    'Confirmez-vous la suppression de l''affectation de la DPI du ',
    ' investi le ',
    'La suppression de l''affectation de la DPI n''est pas autorisée dans ce contexte',
    ' amortie à partir du ',
    ' en ',
    'La suppression de la DPI affectée à l''immobilisation est disponible depuis la fiche immobilisation.',
    'Montant de la DPI     ',
    '          Réintégration obligatoire     ',
    '          Solde     ',
    '                    Ecart',
    'Option invalide, Veuillez valider la sélection des critères.',
    'dont réintégration obligatoire ' );


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 10/05/2006
Modifié le ... :   /  /
Description .. :  Activation de la fenêtre de Synthèse DPI
Mots clefs ... :
*****************************************************************}
Function AMLanceFiche_SyntheseDPI(Stcode: string):String;
begin
   Result:=AGLLanceFiche('AM', 'AMSYNTHESEDPI','' , '', Stcode);
end;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 10/05/2006
Modifié le ... :   /  /
Description .. :  Préparation de la fenêtre DPI
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSYNTHESEDPI.OnArgument (S: String ) ;

Var
  debN  : String;
  finN   : String;
  debN2 : String;
  debN3 : String;
  debN4 : String;
  finN2 : String;
  finN3 : String;
  finN4 : String;
  I,Pos1,Pos2: Integer;
  FDateD : String;
  FDateF : String;

begin

  Inherited ;

  // Récupération des paramêtres d'entrée
  // exercice n-1,n-2,n-3,n-4,n-5,date deb en cours,date fin en cours
  debN := ReadTokenSt(S);
  finN := ReadTokenSt(S);
  FdebN1 := ReadTokenSt(S);
  FfinN1 := ReadTokenSt(S);
  debN2 := ReadTokenSt(S);
  finN2 := ReadTokenSt(S);
  debN3 := ReadTokenSt(S);
  finN3 := ReadTokenSt(S);
  debN4 := ReadTokenSt(S);
  finN4 := ReadTokenSt(S);
  FdebN5 := ReadTokenSt(S);
  FfinN5 := ReadTokenSt(S);
  FDateD := ReadTokenSt(S);
  FDateF := ReadTokenSt(S);

  // Préparation de la fenêtre
  THRadioGroup(GetControl('FS_SRadio')).Onclick := OnClickRadio;
  TToolbarButton97(GetControl('BCherche')).Onclick := OnClickBchercher;
  TToolbarButton97(GetControl('BExport')).Onclick := OnClickBexport;
  TToolbarButton97(GetControl('BImprimer')).Onclick := OnClickBimprimer;
  TToolbarButton97(GetControl('Bdelete')).Onclick := OnClickBSupprimer;
  THValComboBox(GetControl('FS_SCombo')).Onclick := OnClickCombo;
  FListe := THGrid(GetControl('FS_SHGRID'));
  Fradio := THRadioGroup(GetControl('FS_SRadio'));
  FCombo := THValComboBox(GetControl('FS_SCombo'));
  FBvalid := TToolbarButton97(GetControl('BValid'));
  FBsupp := TToolbarButton97(GetControl('Bdelete'));
  FBChercher:= TToolbarButton97(GetControl('BCherche'));
  Ftotal := Thlabel(GetControl('FS_STotal'));
  TFVIERGE(ECRAN).OnKeyDown := OnGrilleKeyDown;

  // Formatage du grid
  Fliste.ColAligns[0]:=taLeftJustify;
  Fliste.ColAligns[1]:=taRightJustify;
  Fliste.ColFormats[1]:=StrfMask(V_PGI.OkDecV, '', True);

  // Sélection par défaut par motif
  Fradio.ItemIndex := 0;

  // Alimentation de la combo
  Fcombo.Items.Add(HM[7]);
  Fcombo.Items.Add(HM[8]+FdebN1+HM[9]+ FfinN1);
  Fcombo.Items.Add(HM[8]+debN2 +HM[9]+ finN2);
  Fcombo.Items.Add(HM[8]+debN3 +HM[9]+ finN3);
  Fcombo.Items.Add(HM[8]+debN4 +HM[9]+ finN4);
  Fcombo.Items.Add(HM[8]+FdebN5+HM[9]+ FfinN5);

  // Présélection 4/09/2006
  I:=0;
  IF FDateF=FfinN1 THEN
    I:=1
  Else IF FDateF=finN2 THEN
    I:=2
  Else IF FDateF=finN3 THEN
    I:=3
  Else IF FDateF=finN4 THEN
    I:=4
  Else IF FDateF=FfinN5 THEN
    I:=5;
  Fcombo.ItemIndex := I;

  // Décalage des boutons TOOLBAR car Bvalider invisible
  pos1 := TToolbarButton97(GetControl('BImprimer')).Left;
  pos2 := TToolbarButton97(GetControl('BValider')).Left ;
  TToolbarButton97(GetControl('BImprimer')).Left := pos2;
  TToolbarButton97(GetControl('BExport')).Left := pos1;

  // Préparation de l'export
  SD := TSaveDialog.Create(Ecran);
  SD.Filter :=
  'Fichier Texte (*.txt)|*.txt|Fichier Excel (*.xls)|*.xls|Fichier Ascii (*.asc)|*.asc|Fichier Lotus (*.wks)|*.wks|Fichier HTML (*.html)|*.html|Fichier XML (*.xml)|*.xml';
  SD.DefaultExt := 'XLS';
  SD.FilterIndex := 1;
  SD.Options := SD.Options + [ofOverwritePrompt, ofPathMustExist,
               ofNoReadonlyReturn, ofNoLongNames] - [ofEnableSizing];

  // Affichage Initial de la fenêtre
  AffichageDeLaListe;

end ;



{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 10/05/2006
Modifié le ... :   /  /
Description .. : Validation de la sélection
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSYNTHESEDPI.OnClickBchercher(Sender: TObject) ;
begin
    AffichageDeLaListe;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 10/05/2006
Modifié le ... :   /  /
Description .. : Modification de la combo
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSYNTHESEDPI.OnClickCombo (Sender: TObject) ;
begin
  // Test modification des paramêtres
  IF (FIndex2 <> FCombo.ItemIndex) Then
     Begin
       FBValid.Visible := False;
       FBChercher.Visible := True;
     End;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 10/05/2006
Modifié le ... :   /  /
Description .. : Modification des radio-boutons
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSYNTHESEDPI.OnClickRadio (Sender: TObject) ;
begin
  // Test modification des paramêtres
  IF (FIndex1 <> Fradio.ItemIndex) Then
     Begin
       FBValid.Visible := False;
       FBChercher.Visible := True;
     End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 10/05/2006
Modifié le ... :   /  /
Description .. : Affichage de la fenêtre
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSYNTHESEDPI.AffichageDeLaListe ;
Var
  Timm,Tims,Tstc,Tcoo,Trei,montant,Total,MT : Double;
  Q,Q1 : Tquery;
  TexteSelection ,Nature,test,immo,dinv,libimmo,cle,Dcle : String;
  svgcol0,svgcol1,SvgCol2,cle1,cle2 : String;
  i,ligne,nbl,compteur1,compteur2  : Integer;
  ok : boolean;
  debeco : TdateTime;
  ddpi: string;
begin

  // Initialisation
  ligne :=0;
  FCadragedroite := '' ;
  cle :='';
  dcle :='';
  Timm:=0;
  Tims:=0;
  Tstc:=0;
  Tcoo:=0;
  Trei:=0;
  Total:=0;

  // --------------------------------------------------------------
  // le thgrid comporte une 3iem colonne invisible (taille à 0)
  // qui permet de stocker la nature et la date d'investissement
  // pour le tri ou la suppression.
  // --------------------------------------------------------------

  // Réinitialisation de la liste
  Nbl := Fliste.rowcount-1 ;
  For i:=1 TO Nbl do
    Begin
      Fliste.CellValues[0,i]:='';
      Fliste.CellValues[1,i]:='';
      Fliste.CellValues[2,i]:='';
    End;
  Fliste.rowcount := 2;

  // Test sélection de l'exercice
  TexteSelection := FCombo.Text;

  IF Copy(TexteSelection,1,4) = 'Tous' Then
     Begin
       // dates entre fin N-5 et fin N-1
       // Les DPI sont saisies à fin d'exercice
       FDatedebut := StrToDate(FfinN5);
       FDateFin   := StrToDate(FfinN1);
     End
  Else
     Begin
       // Date de la sélection
       FDatedebut := StrToDate(copy(TexteSelection,4,10));
       FDateFin   := StrToDate(copy(TexteSelection,18,10));
     End;

  // Lecture de IMMOMVTD
  Q :=OpenSQL('SELECT * FROM IMMOMVTD WHERE IZ_DATE>="'+USDATETIME(FDateDebut)+'" '
                + 'AND IZ_DATE<="'+USDATETIME(FDateFin)+'"'
                + 'AND IZ_NATURED<>"DPI"' , FALSE) ;
  try
     While Not Q.Eof do
        begin

          Nature := Q.FindField('IZ_NATURED').AsString;
          Montant:= Q.FindField('IZ_Montant').AsFloat;
          Immo   := Q.FindField('IZ_IMMO').AsString; ;
          Dinv   := Copy(Q.FindField('IZ_LIBELLE').AsString,1,10);
          Dcle   := Copy(Dinv,7,4)+ Copy(Dinv,4,2)+ Copy(Dinv,1,2);
          Total  := Total + montant;
          ddpi   := Q.FindField('IZ_DATE').AsString;

          // Groupées
          if (GetControlText('FS_SRadio')= '1') then
             Begin
               IF Nature = 'IMM' Then
                  Timm := Timm + Montant
               Else IF Nature = 'IMS' Then
                  Tims := Tims + Montant
               Else IF Nature = 'STC' Then
                  Tstc := Tstc + Montant
               Else IF Nature = 'COO' Then
                  Tcoo := Tcoo + Montant
               Else IF Nature = 'REI' Then
                  Trei := Trei + Montant
             end

          // Dégroupées
          else if (GetControlText('FS_SRadio')= '2') then

             Begin
               IF Nature = 'IMM' Then
                  Begin
                     // Lecture du libellé
                     Q1:=OPENSQL('SELECT I_LIBELLE,I_DATEDEBECO FROM IMMO WHERE I_IMMO ="'+immo+'"',True);
                     If Not Q1.EOF Then
                        Begin
                          Libimmo:=Q1.FindField('I_LIBELLE').AsString;
                          Debeco :=Q1.FindField('I_DATEDEBECO').AsDateTime;
                        End
                     Else
                        Debeco := iDate1900;
                     Ferme(Q1);

                     // Présentation détaillées des immos affectés
                     ligne := ligne+1 ;
                     cle := '1        ';
                     //Fliste.CellValues[0,ligne]:=HM[5]+' '+immo+' '+libimmo;
                     Fliste.CellValues[0,ligne]:=HM[5]+' '+libimmo+ HM[20]+ DateToStr(Debeco) +' ('+immo+')';
                     Fliste.CellValues[1,ligne]:=StrFMontant(montant,15,V_PGI.OkDecV,'',True);
                     Fliste.CellValues[2,ligne]:=cle + immo ;
                   End
               Else IF Nature = 'IMS' Then
                   Begin
                     // Présentation détaillées des immos sorties
                     ligne := ligne+1 ;
                     cle := '2'+Dcle ;
                     Fliste.CellValues[0,ligne]:=HM[6]+ HM[14] +Dinv+', DPI du '+ddpi+')';
                     Fliste.CellValues[1,ligne]:=StrFMontant(montant,15,V_PGI.OkDecV,'',True);
                     Fliste.CellValues[2,ligne]:=cle;
                   End
               Else IF Nature = 'STC' Then
                   Begin
                     // Présentation détaillées des stocks
                     ligne := ligne+1 ;
                     cle := '3'+Dcle ;
                     Fliste.CellValues[0,ligne]:=HM[1]+ HM[14] +Dinv+', DPI du '+ddpi+')';
                     Fliste.CellValues[1,ligne]:=StrFMontant(montant,15,V_PGI.OkDecV,'',True);
                     Fliste.CellValues[2,ligne]:=cle;
                   End
               Else IF Nature = 'COO' Then
                   Begin
                     // Présentation détaillées des parts de coopératives
                     ligne := ligne+1 ;
                     cle := '4'+Dcle ;
                     Fliste.CellValues[0,ligne]:=HM[2]+ HM[14] +Dinv+', DPI du '+ddpi+')';
                     Fliste.CellValues[1,ligne]:=StrFMontant(montant,15,V_PGI.OkDecV,'',True);
                     Fliste.CellValues[2,ligne]:=cle;
                   End
               Else IF Nature = 'REI' Then
                   Begin
                     // Présentation détaillées des réinvestissemts
                     ligne := ligne+1 ;
                     cle := '5'+Dcle ;
                     Fliste.CellValues[0,ligne]:=HM[3]+ HM[15] +Dinv + ' (DPI du '+ddpi+')' ;
                     Fliste.CellValues[1,ligne]:=StrFMontant(montant,15,V_PGI.OkDecV,'',True);
                     Fliste.CellValues[2,ligne]:=cle;
                   End
             end


          Else
             // Par exercice
             Begin
                // Recherche dans le grid si enreg présent
                ok :=False ;
                For i:=1 TO ligne do
                  Begin
                     test := Copy(Fliste.CellValues[0,i],1,10);
                     IF test = Dinv Then
                        Begin
                          // Mise à jour
                          Mt:= valeur(Fliste.CellValues[1,i])+ Montant;
                          Fliste.CellValues[1,i]:= StrFMontant(Mt,15,V_PGI.OkDecV,'',True);
                          ok:=True;
                          Break;
                        end;
                  end;

                IF ok = False Then
                  // Création
                  Begin
                     ligne := ligne+1 ;
                     Fliste.CellValues[0,ligne]:= Dinv;
                     Fliste.CellValues[1,ligne]:=StrFMontant(montant,15,V_PGI.OkDecV,'',True);
                     Fliste.CellValues[2,ligne]:='1'+Dcle;
                  End;
             End;

          Q.Next ;

        End;
  finally
     Ferme(Q);
  End;

  // Groupées
  if (GetControlText('FS_SRadio')= '1') Then
    Begin
      // Affichage des totaux
      Fliste.CellValues[0,1]:=HM[0];
      Fliste.CellValues[1,1]:=StrFMontant(Timm+Tims,15,V_PGI.OkDecV,'',True);
      Fliste.CellValues[0,2]:=HM[1];
      Fliste.CellValues[1,2]:=StrFMontant(Tstc,15,V_PGI.OkDecV,'',True);
      Fliste.CellValues[0,3]:=HM[2];
      Fliste.CellValues[1,3]:=StrFMontant(Tcoo,15,V_PGI.OkDecV,'',True);
      Fliste.CellValues[0,4]:=HM[3];
      Fliste.CellValues[1,4]:=StrFMontant(Trei,15,V_PGI.OkDecV,'',True);
      Fliste.RowCount:=5;
    End

  // Dégroupées ou exercice tri sur la date
  Else
    Begin
      // pas de tri si 1 ligne
      IF ligne>1 Then
         Begin
            // Tri sur la clé (1 à 4 + date investissement)
           For compteur1 := 1 to ligne do
             Begin
               ok:=False;
               While ok=False Do
                  Begin
                    Cle1 := Copy(Fliste.CellValues[2,compteur1],1,9);
                    For compteur2 := compteur1 + 1 to ligne Do
                      Begin
                        Cle2 := Copy(Fliste.CellValues[2,compteur2],1,9);
                        IF Cle1 > Cle2 THEN
                          Begin
                            // Svg ligne 1
                            SvgCol0 := Fliste.CellValues[0,compteur1];
                            SvgCol1 := Fliste.CellValues[1,compteur1];
                            SvgCol2 := Fliste.CellValues[2,compteur1];
                            // Ligne 1 = ligne2
                            Fliste.CellValues[0,compteur1] :=Fliste.CellValues[0,compteur2];
                            Fliste.CellValues[1,compteur1] :=Fliste.CellValues[1,compteur2];
                            Fliste.CellValues[2,compteur1] :=Fliste.CellValues[2,compteur2];
                            // Ligne 2 = ligne1
                            Fliste.CellValues[0,compteur2] :=SvgCol0;
                            Fliste.CellValues[1,compteur2] :=SvgCol1;
                            Fliste.CellValues[2,compteur2] :=SvgCol2;
                            // sortie de la boucle
                            break;
                          End;
                      End;

                    IF (compteur2-1) = ligne Then
                       Ok:=True;

                  End;
             End;

         // Détermination du nombre de ligne
         Fliste.RowCount:=ligne+1;

         End;
    End;

    // Si nombre de ligne >8 ==> ascenseur
    IF Fliste.RowCount>8 Then
       Begin
          FCadragedroite := '       ' ;
          For compteur1 :=1 TO ligne Do
             Fliste.CellValues[1,Compteur1]:= Fliste.CellValues[1,Compteur1] + FCadragedroite ;
        End;

  // Affichage du total
  SetControlCaption('FS_STotal',StrFMontant(total,15,V_PGI.OkDecV,'',True)+FCadragedroite);

  // Affichage de la réintégration obligatoire (n-5) ou du solde (n-1 à n-4)
  AffichageDuSolde;

  //Focus sur le grid
  Fliste.SetFocus;

  // Sauvegarde des options de traitement
  FIndex1 := Fradio.ItemIndex;
  FIndex2 := FCombo.ItemIndex;

  // Positionnement des boutons
  FBValid.Visible := True;
  FBChercher.Visible := False;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 10/05/2006
Modifié le ... :   /  /
Description .. : Affichage du solde
Mots clefs ... :
*****************************************************************}
Procedure TOF_AMSYNTHESEDPI.AffichageDuSolde ;
Var
  Q2 : TQuery;
  montant,montant1,total,MDPI,Solde : Double;
begin

  MDPI:=0;
  MONTANT:=0;
  MONTANT1:=0;

  // Affichage de la réintégration obligatoire (n-5) ou du solde (n-1 à n-4)
  If Fcombo.ItemIndex <>0 Then
    Begin
      // Lecture de la dpi sur l'année
      Q2 :=OpenSQL('SELECT * FROM IMMOMVTD WHERE IZ_DATE="'+USDATETIME(FDateFin)+'" '
             + 'AND IZ_NATURED="DPI"', FALSE) ;
      If Not Q2.Eof  Then
         MDPI:= Q2.FindField('IZ_MONTANT').AsFloat;
      Ferme(Q2);
    End
  Else
    Begin
      // Lecture de la dpi sur tous les exercices
      Q2 :=OpenSQL('SELECT * FROM IMMOMVTD WHERE IZ_DATE>="'+USDATETIME(StrToDate(FfinN5))+'" '
             + 'AND IZ_DATE<="'+USDATETIME(StrToDate(FfinN1))+'"'
             + 'AND IZ_NATURED="DPI"' , FALSE) ;
      Try
        While Not Q2.Eof Do
          Begin
            MDPI:= MDPI + Q2.FindField('IZ_MONTANT').AsFloat;
            Q2.next;
          End;
      Finally
        Ferme(Q2);
      End;
    End;

  // Calcul du solde
  total := Valeur(GetControlText('FS_STOTAL'));
  Solde := MDPI - total;

  // Affichage des montants
  SetControlCaption('FS_SDPI',StrFMontant(MDPI,15,V_PGI.OkDecV,'',True));
  SetControlCaption('FS_SSolde',StrFMontant(Solde,15,V_PGI.OkDecV,'',True));

  // Affichage de la dpi
  THLabel(GetControl('FS_Slabel3')).visible :=True;
  THLabel(GetControl('FS_SDPI')).visible :=True;

  // cache dont réintégration
  THLabel(GetControl('FS_Slabel1')).visible :=False;
  THLabel(GetControl('FS_Reint')).visible :=False;

  // Si N-5 modification des libellés
  If Fcombo.ItemIndex = 5 Then
    Begin
      // Affichage de la réintégration
      THLabel(GetControl('FS_Slabel4')).visible :=True;
      THLabel(GetControl('FS_Slabel5')).visible :=False;
      THLabel(GetControl('FS_SSolde')).left := 370;
    End
  Else
    Begin
      // Affichage du solde si <> n-5
      THLabel(GetControl('FS_Slabel4')).visible :=False;
      THLabel(GetControl('FS_Slabel5')).visible :=True;
      THLabel(GetControl('FS_SSolde')).left := 267;

      //  Si tous les exercices
      If Fcombo.ItemIndex = 0 Then
        Begin
          THLabel(GetControl('FS_Slabel1')).visible :=True;
          THLabel(GetControl('FS_Reint')).visible :=True;

          // Lecture des affectations sur N-5
          Q2 :=OpenSQL('SELECT * FROM IMMOMVTD WHERE IZ_DATE>="'+USDATETIME(StrtoDate(FdebN5))+'" '
                  + 'AND IZ_DATE<="'+USDATETIME(StrToDate(FfinN5))+'"', FALSE) ;
          Try
            While Not Q2.Eof do
              begin
                IF Q2.FindField('IZ_NATURED').AsString <>'DPI' Then
                  Montant:=Montant+ Q2.FindField('IZ_Montant').AsFloat
                ELSE
                  Montant1:=Montant1+ Q2.FindField('IZ_Montant').AsFloat;
                Q2.Next;
              End;
          Finally
            Ferme(Q2)
          End;
          Montant1 := Montant1 - Montant;
          SetControlCaption('FS_REINT',StrFMontant(Montant1,15,V_PGI.OkDecV,'',True));
          // Pour aligner avec le solde
          // MVG 16/10/2006 FQ 18951
//          THLabel(GetControl('FS_SSolde')).left := THLabel(GetControl('FS_REINT')).left; // MVG 390;
        End;
    End;

End;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 10/05/2006
Modifié le ... :   /  /
Description .. : Validation des critères sur Touche F9
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSYNTHESEDPI.OnGrilleKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    VK_F9 : // Revalidation de la fenêtre
      AffichageDeLaListe;
    46:{Ctrl+Suppr}
       IF Shift=[ssCtrl] Then
         OnClickBSupprimer(Sender);
    end;
end;



{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 10/05/2006
Modifié le ... :   /  /
Description .. : Impression du grid
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSYNTHESEDPI.OnClickBImprimer(Sender: TObject);
Var
  PrintCaption : String ;
  Edate : String;
begin
  IF (FBValid.Visible = True) Then
     Begin
       PrintCaption := '';
       Edate := ' ('+DateToStr(FDateDebut)+' - '+ DateToStr(FDateFin)+')';
       PreparePrintGrid;

       // Mise à jour du  titre
       if (GetControlText('FS_SRadio')= '1') Then
          PrintCaption:=PrintCaption+HM[10]+ Edate
       Else if (GetControlText('FS_SRadio')= '2') Then
          PrintCaption:=PrintCaption+HM[11]+ Edate
       Else
          PrintCaption:=PrintCaption+HM[12]+ Edate;

       {$IFDEF EAGLCLIENT}
          PrintDBGrid( PrintCaption, FListe.ListeParam, '', '');
       {$ELSE}
          PrintDBGrid(Fliste, nil, PrintCaption, '') ;
       {$ENDIF}

        // Suppression de la ligne total
        Fliste.RowCount:= Fliste.RowCount-1;

       If Fcombo.ItemIndex = 0 Then
          // Suppression de la ligne réintégration
          Fliste.RowCount:= Fliste.RowCount-1;

     End
  Else
     Pgibox(HM[27]);

end;



{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 10/05/2006
Modifié le ... :   /  /
Description .. : Suppression d'une affectation
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSYNTHESEDPI.OnClickBsupprimer (Sender: TObject) ;
Var
  i,code,nat : Integer;
  dinv,cle,nature : String;
  question : TmodalResult;
  montant,total : double;
  supp : Boolean;
begin

  supp:=false;
  // Activation du bouton supprimer une affectation si détaillées et dans 1 exercice
  IF ((FBValid.Visible = False) OR
     (GetControlText('FS_SRadio')<> '2') OR
     (Copy((FCombo.Text),1,4) = 'Tous')) Then
     PgiBox (HM[19])
  Else
    Begin

       // Détermination dee éléments de la ligne sélectionnée
      i:=Fliste.row;

      IF (i=1) AND (Fliste.CellValues[0,i]='') Then
        // Pas de ligne
      ELSE
        Begin
          cle := Fliste.CellValues[2,i];
          dinv := Copy(Fliste.CellValues[2,i],2,8);
          dinv := Copy(Dinv,7,2)+'/'+Copy(Dinv,5,2)+'/'+Copy(Dinv,1,4);
          code := StrToInt(copy(cle,1,1));
          montant := Valeur(trim(Fliste.CellValues[1,i]));

          // L'affectation en cours est une immobilisation
          IF code =1 Then
             // 25/06/2006 sup sur immo interdite
             PgiBox (HM[22])
          ELSE
             Begin
               nat:=0;
               case code of
                 2:  begin nat := 6; nature:='IMS'; End;
                 3:  begin nat := 1; nature:='STC'; End;
                 4:  begin nat := 2; nature:='COO'; End;
                 5:  begin nat := 3; nature:='REI'; End;
               end;
               Question := PGIAsk(HM[17]+DateToStr(Fdatedebut)+HM[9]+DateToStr(FdateFin)+ chr(13)
                           +'('+HM[18]+dinv+HM[21]+LowerCase(HM[nat])+') ?.',Ecran.Caption);
               If Question = mrYes then
                  Begin
                    ExecuteSQL ('DELETE FROM IMMOMVTD WHERE IZ_DATE="'+USDATETIME(FDateFin)+'"'
                               + 'AND IZ_NATURED ="'+nature+'"' );
                    supp:=true;
                  End;
             End;

          // Supression de la ligne dans la liste
          IF supp=True Then
             Begin
               // Affichage du total
               total := Valeur(GetControlText('FS_STOTAL'))-montant;
               SetControlCaption('FS_STotal',StrFMontant(total,15,V_PGI.OkDecV,'',True)+FCadragedroite);
               Fliste.DeleteRow(i);
               Fliste.refresh;
               AffichageDuSolde;
             End;
        end;
    end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 10/05/2006
Modifié le ... :   /  /
Description .. : Export du grid
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSYNTHESEDPI.OnClickBExport(Sender: TObject);

begin
  IF (FBValid.Visible = True) Then
     Begin
       PreparePrintGrid;
       if SD.Execute then
         ExportGrid(Fliste, nil, SD.FileName, SD.FilterIndex, TRUE) ;

       // Suppression de la ligne total
       Fliste.RowCount:= Fliste.RowCount-1;

       If Fcombo.ItemIndex = 0 Then
          // Suppression de la ligne réintégration
          Fliste.RowCount:= Fliste.RowCount-1;

     End
  Else
     Pgibox(HM[27]);
end;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 10/05/2006
Modifié le ... :   /  /
Description .. : Préparation du grid pour impression ou export
Mots clefs ... :
*****************************************************************}
procedure  TOF_AMSYNTHESEDPI.PreparePrintGrid;
var
   ligne : integer;
   montant : Double;
   texte : string;
begin
   // Ajout de la ligne total
   Ligne := Fliste.RowCount;

   // montant de la DPI
   Texte := HM[23] + GetControlText('FS_SDPI');
   // Reintégration ou solde
   If Fcombo.ItemIndex = 5 Then
     Texte := Texte + HM[24]
   Else
     Texte := Texte + HM[25];
   Texte := Texte + GetControlText('FS_SSolde');

   // Total
   If Fcombo.ItemIndex <> 5 Then
     Texte := Texte + '                    ';

   Texte := Texte + HM[26];
   Fliste.CellValues[0,ligne]:=Texte;

   montant := Valeur(GetControlText('FS_STOTAL'));
   Fliste.CellValues[1,ligne]:=StrFMontant(Montant,15,V_PGI.OkDecV,'',True)+FCadragedroite;

   Fliste.RowCount:=ligne+1;

   If Fcombo.ItemIndex = 0 Then
     Begin
       ligne:=ligne+1;
       Fliste.CellValues[0,ligne]:=HM[28]+'      '+GetControlText('FS_REINT');
       Fliste.RowCount:=ligne+1;
     End;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 10/05/2006
Modifié le ... :   /  /
Description .. : Fermeture de la synthèse
Mots clefs ... :
*****************************************************************}
procedure  TOF_AMSYNTHESEDPI.OnClose;
begin
  FListe.Free;
  if assigned(SD)then
    SD.Free;
end;



{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 26/06/2006
Modifié le ... :   /  /
Description .. : Fontion public, maj de immomvtd suite à la suppression d'une immo

Paramêtre d'entrée :

   Action                nimmo1                    nimmo2                     fraction
  ---------------------------------------------------------------------------------------------
  S =suppression        immo à supprimer        rien                          rien
  M =Mutation           immo d'origine          nelle immo mutée              rien
  E =Eclatement         immo à éclater (Mère)   nelle immo éclatée (fille)    % du mt à éclater
  W =Supp d'eclatement  immo d'origine (Mère)   immo éclatée à supp (fille)   rien

suite .......... : modif mbo le 17.10.06 pour pb d'arrondi
*****************************************************************}
Procedure AM_MAJ_IMMOMVTD(action: String ; nimmo1:String ; nimmo2:String ; Fraction:double);
Var
  sequence : integer;
  montant, MontantMere  : double;
  datetest : TdateTime;
  Tlog : Tob;
  Q,Q1 : Tquery;
Begin

  IF action = 'S' Then

     // Si suppression d'immo suppression des enregs dans IMMOMVTD
     ExecuteSQL ('DELETE FROM IMMOMVTD WHERE IZ_IMMO="'+nimmo1+'" AND IZ_NATURED="IMM"')

  Else If action = 'M' Then

     // Si mutation d'immo, maj des enregs dans IMMOMVTD
     ExecuteSQL ('UPDATE IMMOMVTD SET IZ_IMMO="'+nimmo2+'" WHERE IZ_IMMO="'+nimmo1+'"'
                 + ' AND IZ_NATURED="IMM"')

  Else If action = 'E' Then
     Begin
        // Si éclatement d'immo, cr&ation et maj des enregs dans IMMOMVTD
        // Lecture des immos d'origine dans IMMOMVTD
        Q :=OpenSQL('SELECT * FROM IMMOMVTD WHERE IZ_IMMO="'+nimmo1+'" AND IZ_NATURED="IMM"',False);
        Try
           While Not Q.Eof do
              Begin
                 // Recherche dernier numéro de séquence pour la date d(investissement
                 DateTest := Q.FindField('IZ_DATE').AsDateTime;
                 sequence := 0 ;

                 Q1 :=OpenSQL('SELECT * FROM IMMOMVTD WHERE IZ_DATE="'+USDATETIME(DateTest)+'"', FALSE) ;
                 try
                    While Not Q1.Eof do
                       Begin
                         IF Q1.FindField('IZ_SEQUENCE').AsINTEGER > sequence THEN
                            sequence := Q1.FindField('IZ_SEQUENCE').AsINTEGER;
                         Q1.Next ;
                       End;
                 finally
                    Ferme(Q1);
                 End;

                 // Création complet pour l'immo éclaté

                 // calcul du montant éclaté
                 Montant:= Q.FindField('IZ_MONTANT').AsFloat * fraction ;

                 TLog := TOB.Create ('IMMOMVTD',nil,-1);
                 try
                    TLog.PutValue('IZ_SEQUENCE', sequence + 1) ;
                    TLog.PutValue('IZ_DATE', Q.FindField('IZ_DATE').AsDateTime );
                    TLog.PutValue('IZ_MONTANT',montant) ;
                    TLog.PutValue('IZ_NATURED','IMM') ;
                    TLog.PutValue('IZ_LIBELLE',Q.FindField('IZ_LIBELLE').AsString) ;
                    TLog.PutValue('IZ_IMMO', nimmo2) ;
                    TLog.PutValue('IZ_LIBELLE',Q.FindField('IZ_LIBELLE').AsString) ;
                    TLog.InsertDB(nil);
                 finally
                    TLog.Free;
                 End;

                 // Mise à jour du montant de l'immo d'origine
                 Q.Edit;
                 // Montant total à éclaté
                 // modif mbo 17.10.06 pour des pb d'arrondi : on enlève de la mère le montant de la fille
                 // Montant:= Q.FindField('IZ_MONTANT').AsFloat * (1-fraction) ;
                 //Q.FindField('IZ_MONTANT').AsFloat := montant ;
                 MontantMere:= Q.FindField('IZ_MONTANT').AsFloat - montant ;
                 Q.FindField('IZ_MONTANT').AsFloat := MontantMere ;

                 Q.Post;
                 // immo suivante
                 Q.Next;
              End;
        Finally
           Ferme(Q);
        End;
     End

  Else If action = 'W' Then

     Begin
        // Si suppression d'éclatement d'immo
        // Lecture des immos (mère) d'origine dans IMMOMVTD
        Q :=OpenSQL('SELECT * FROM IMMOMVTD WHERE IZ_IMMO="'+nimmo1+'" AND IZ_NATURED="IMM"',False);
        Try
           While Not Q.Eof do
              Begin

                 // Recherche de l'immo éclatée (fille) pour récupérer le montant
                 DateTest := Q.FindField('IZ_DATE').AsDateTime;
                 Montant := 0 ;
                 Q1 :=OpenSQL('SELECT * FROM IMMOMVTD WHERE IZ_DATE="'+USDATETIME(DateTest)+'"'
                                 + ' AND IZ_IMMO="'+nimmo2+'"', FALSE) ;
                 If Not Q1.Eof  Then
                    Montant:= Q1.FindField('IZ_MONTANT').AsFloat;
                 Ferme(Q1);

                 // Mise à jour du montant de l'immo mère
                 Q.Edit;
                 // Montant total à éclaté
                 Q.FindField('IZ_MONTANT').AsFloat := montant + Q.FindField('IZ_MONTANT').AsFloat ;
                 Q.Post;

                 // suppression de l'enreg dans IMMOMVTD de la fille
                 ExecuteSQL ('DELETE FROM IMMOMVTD WHERE IZ_IMMO="'+nimmo2+'"'
                                  + 'AND IZ_DATE="'+USDATETIME(DateTest)+'"');

                 // suivant
                 Q.next;

              End;
        Finally
           Ferme(Q);
        End;
     End;


End;





Initialization
  registerclasses ( [ TOF_AMSYNTHESEDPI ] ) ;
end.
