{***********UNITE*************************************************
Auteur  ...... : Paie Pgi  RMA
Créé le ...... : 19/04/2006
Modifié le ... :   /  /
Description .. : Suivi des effectifs mensuellements répartis par code stat
Suite ........ :    (TRAVAILNx, STAT, LIBREx)
Mots clefs ... :
*****************************************************************
PT01 RMA 15/10/07 FQ14792-FQ14317 Pb sur paie décalée + comptages erronés
}
Unit UTOFSUIVIEFFECTIFS;

Interface

Uses StdCtrls, Classes, sysutils, ComCtrls,
  {$IFDEF EAGLCLIENT}
     eQRS1,
     UtileAgl,
  {$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     QRS1,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF,
  ParamSoc, HQry,
  UTOB ,UtobDebug ;

Type
  TOF_SUIVI_EFFECTIFS = Class (TOF)
    private
      OnSort  : Boolean;   //True ==> On a fait la X pour sortir ,on ne test plus la validé de la période saisie
      TobEdit : Tob;
      CodeSel : String;

      Procedure Change(Sender: TObject) ;
      Procedure VerifPeriode1(Sender: TObject) ;
      Procedure VerifPeriode2(Sender: TObject) ;
      Procedure ControleChampCompl ;
      Procedure ChangeExercice(Sender: TObject);
      Procedure TraitChangeExercice ;
      Procedure ControleCheckRupture ;
    public
      Procedure OnArgument (S : String ) ; override ;
      Procedure OnUpdate                 ; override ;
      Procedure OnClose                  ; override ;
   end ;

Implementation
Uses EntPaie;

Procedure TOF_SUIVI_EFFECTIFS.OnArgument (S : String ) ;
var
  Defaut   : THEdit;
  Exercice : THValComboBox ;
  Check    : TCheckBox;
  SQL, N_an: String;
  I : Integer;
  Q :TQuery;

Begin
  Inherited ;
  OnSort := False;
  CodeSel:= '';
  TFQRS1(Ecran).Caption := 'Suivi des effectifs';
  TFQRS1(Ecran).CodeEtat:= 'PSE';
  UpdateCaption(Ecran);

  Defaut := ThEdit(Getcontrol('DOSSIER'));
  If Defaut <> nil Then Defaut.text := GetParamSoc('SO_LIBELLE');
  SetControlChecked('PAIEDECALEE',VH_Paie.PGDecalage);
  N_an := '';

  SQL := 'SELECT MAX(PEX_EXERCICE) FROM EXERSOCIAL WHERE PEX_ACTIF="X"';
  Q := OpenSQL(SQL,True);

  If not Q.eof Then
  Begin
    N_an := Q.Fields[0].asstring;
  End;
  Ferme(Q);
  If N_an <> '' Then SetControlText('EDEXERSOC',N_an);
  TraitChangeExercice;

  Exercice:= THValComboBox(getcontrol('EDEXERSOC'));
  If Exercice <> nil Then Exercice.OnChange := ChangeExercice;
  Defaut := THEdit(GetControl('XX_VARIABLEDEB'));
  If Defaut <> nil Then Defaut.OnExit := VerifPeriode1;
  Defaut := THEdit(GetControl('XX_VARIABLEFIN'));
  If Defaut <> nil Then Defaut.OnExit := VerifPeriode2;

  ControleChampCompl ;
  SetControlProperty('TBCOMPLEMENT', 'Tabvisible', (VH_Paie.PGNbreStatOrg > 0) or (VH_Paie.PGLibCodeStat <> ''));
  SetControlProperty('TBCHAMPLIBRE', 'Tabvisible', (VH_Paie.PgNbCombo > 0));

  For I := 1 to 5 do
  Begin
    Check := TCheckBox(GetControl('CN'+IntToStr(I)));
    If Check <> nil Then Check.OnClick := Change;
  End;
  For I := 1 to 4 do
  Begin
    Check := TCheckBox(GetControl('CL'+IntToStr(I)));
    If Check <> nil Then Check.OnClick := Change;
  End;
End;

Procedure TOF_SUIVI_EFFECTIFS.OnUpdate ;
var
  I ,Y ,MM ,NonMouv : Integer;
  Check : TCheckBox;
  TH_Titre, TH_Champ :THEdit;
  DatDeb, DatFin     :TDateTime;
  Pages :TPageControl;
  SQL, AjoutChamp, Where, Suite, Champ : String;
  Q :TQuery;
  TSQL ,Tetab ,Ttot ,TFille ,Ttrav : Tob;

Begin
  Inherited ;
  Pages := TPageControl(GetControl('Pages'));
  Where := RecupWhereCritere(Pages);
  TH_Titre :=THEdit(GetControl('TITRE'));
  TH_Champ :=THEdit(GetControl('CHAMP'));
  DatDeb := StrToDate(GetControltext('XX_VARIABLEDEB'));
  DatFin := StrToDate(GetControltext('XX_VARIABLEFIN'));
  AjoutChamp :='';
  Suite := '';
  NonMouv := 0;

  If (TH_Titre <> nil) And (TH_Champ <> nil) Then
  Begin
    If Trim(TH_Champ.Text)<>'' Then AjoutChamp := Trim(TH_Champ.Text) + ' As CSel,' ;
    If Trim(Where)= '' Then Where := 'WHERE'
       Else Suite := ' AND';
  End;
  SQL:='SELECT COUNT(PSA_SALARIE) FROM SALARIES Where PSA_SALARIE ' +
       'NOT IN (SELECT DISTINCT PHC_SALARIE From HISTOCUMSAL Where ' +
       '(PHC_DATEDEBUT >= "'+UsDateTime(DatDeb)+'" AND PHC_DATEFIN <= "'+UsDateTime(DatFin)+'")) ' +
       'And ((PSA_DATEENTREE <= "'+UsDateTime(DatFin)+'") And ' +
       '(PSA_DATESORTIE >= "'+UsDateTime(DatDeb)+'" Or PSA_DATESORTIE = "01/01/1900")) ';

  Q := OpenSQL(SQL,True);

  If not Q.eof Then
  Begin
    NonMouv := Q.Fields[0].asInteger;
  End;

  FreeAndNil (TobEdit);
  TobEdit := Tob.Create('Suivi_Effectif',Nil,-1);
  SQL :='SELECT Distinct PHC_ETABLISSEMENT,ET_LIBELLE,PHC_SALARIE,MONTH(PHC_DATEFIN) MOIS ,'+ AjoutChamp + 'PSA_DATESORTIE' +
        ' FROM HISTOCUMSAL' +
        ' LEFT JOIN ETABLISS ON PHC_ETABLISSEMENT = ET_ETABLISSEMENT' +
        ' LEFT JOIN SALARIES ON PHC_SALARIE = PSA_SALARIE ';
  //PT01 SQL := SQL + Where + Suite + ' (PHC_CUMULPAIE = "01") And ' +
  //PT01     '(PHC_DATEDEBUT >= "'+UsDateTime(DatDeb)+'" AND PHC_DATEFIN <= "'+UsDateTime(DatFin)+'") ';
  SQL := SQL + Where + Suite + ' (PHC_REPRISE <> "X") And ' +
      '(PHC_DATEDEBUT >= "'+UsDateTime(DatDeb)+'" AND PHC_DATEFIN <= "'+UsDateTime(DatFin)+'") ';

  Check := TCheckBox(GetControl('FINPERIODE'));
  If (Check <> nil) Then
    If Check.Checked = True then
       SQL := SQL + 'AND (PSA_DATESORTIE <= "'+UsDateTime(IDate1900)+'"'+
                        ' OR PSA_DATESORTIE IS NULL '+
                        ' OR PSA_DATESORTIE >= "'+UsDateTime(DatFin)+'") ';

  If CodeSel <> '' Then
      SQL := SQL + 'ORDER BY PHC_ETABLISSEMENT,ET_LIBELLE,PHC_SALARIE,MOIS,CSEL,PSA_DATESORTIE'
    Else
      SQL := SQL + 'ORDER BY PHC_ETABLISSEMENT,ET_LIBELLE,PHC_SALARIE,MOIS,PSA_DATESORTIE';

  Q := OpenSQL(SQL,True);
  TSQL := Tob.Create('Donnees',Nil,-1);
  TSQL.LoadDetailDB('Donnees','','',Q,False);
  Ferme(Q);

  Tetab := Tob.Create('Tot_Etab',Nil,-1);
  Ttot  := Tob.Create('Tot_Dossier',Nil,-1);

  If CodeSel <> '' Then
  Begin
     SQL := 'SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD Where CC_TYPE = "'+CodeSel+'"';
     Q := OpenSQL(SQL,True);
     Ttrav := Tob.Create('Tob_Travail',Nil,-1);
     Ttrav.LoadDetailDB('Tob_Travail','','',Q,False);
     Ferme(Q);

     For i := 0 To Ttrav.Detail.Count -1 do
     Begin
        TFille := Tob.Create('Fille',Tetab,-1);
        TFille.AddChampSupValeur('CC_CODE',Ttrav.Detail[i].GetValue('CC_CODE'));
        TFille.AddChampSupValeur('CC_LIBELLE',Ttrav.Detail[i].GetValue('CC_LIBELLE'));
        TFille.AddChampSupValeur('M01',0);
        TFille.AddChampSupValeur('M02',0);
        TFille.AddChampSupValeur('M03',0);
        TFille.AddChampSupValeur('M04',0);
        TFille.AddChampSupValeur('M05',0);
        TFille.AddChampSupValeur('M06',0);
        TFille.AddChampSupValeur('M07',0);
        TFille.AddChampSupValeur('M08',0);
        TFille.AddChampSupValeur('M09',0);
        TFille.AddChampSupValeur('M10',0);
        TFille.AddChampSupValeur('M11',0);
        TFille.AddChampSupValeur('M12',0);

        TFille := Tob.Create('Fille',Ttot,-1);
        TFille.AddChampSupValeur('CC_CODE',Ttrav.Detail[i].GetValue('CC_CODE'));
        TFille.AddChampSupValeur('CC_LIBELLE',Ttrav.Detail[i].GetValue('CC_LIBELLE'));
        TFille.AddChampSupValeur('M01',0);
        TFille.AddChampSupValeur('M02',0);
        TFille.AddChampSupValeur('M03',0);
        TFille.AddChampSupValeur('M04',0);
        TFille.AddChampSupValeur('M05',0);
        TFille.AddChampSupValeur('M06',0);
        TFille.AddChampSupValeur('M07',0);
        TFille.AddChampSupValeur('M08',0);
        TFille.AddChampSupValeur('M09',0);
        TFille.AddChampSupValeur('M10',0);
        TFille.AddChampSupValeur('M11',0);
        TFille.AddChampSupValeur('M12',0);
     End;
     Ttrav.Free;
  End;
  TFille := Tob.Create('Fille',Tetab,-1);
  TFille.AddChampSupValeur('CC_CODE','XXX');
  TFille.AddChampSupValeur('CC_LIBELLE','Sans affectation');
  TFille.AddChampSupValeur('M01',0);
  TFille.AddChampSupValeur('M02',0);
  TFille.AddChampSupValeur('M03',0);
  TFille.AddChampSupValeur('M04',0);
  TFille.AddChampSupValeur('M05',0);
  TFille.AddChampSupValeur('M06',0);
  TFille.AddChampSupValeur('M07',0);
  TFille.AddChampSupValeur('M08',0);
  TFille.AddChampSupValeur('M09',0);
  TFille.AddChampSupValeur('M10',0);
  TFille.AddChampSupValeur('M11',0);
  TFille.AddChampSupValeur('M12',0);

  TFille := Tob.Create('Fille',Ttot,-1);
  TFille.AddChampSupValeur('CC_CODE','XXX');
  TFille.AddChampSupValeur('CC_LIBELLE','Sans affectation');
  TFille.AddChampSupValeur('M01',0);
  TFille.AddChampSupValeur('M02',0);
  TFille.AddChampSupValeur('M03',0);
  TFille.AddChampSupValeur('M04',0);
  TFille.AddChampSupValeur('M05',0);
  TFille.AddChampSupValeur('M06',0);
  TFille.AddChampSupValeur('M07',0);
  TFille.AddChampSupValeur('M08',0);
  TFille.AddChampSupValeur('M09',0);
  TFille.AddChampSupValeur('M10',0);
  TFille.AddChampSupValeur('M11',0);
  TFille.AddChampSupValeur('M12',0);

  For i := 0 To TSQL.Detail.Count -1 do
  begin
     y := 0;
     If CodeSel <> '' Then
     Begin
        For y := 0 To Tetab.Detail.Count -1 do
           If Tetab.Detail[y].GetValue('CC_CODE') = TSQL.Detail[i].GetValue('CSel') Then Break;
     End;
     If (y > Tetab.Detail.Count -1) Or (CodeSel = '') Then y := Tetab.Detail.Count -1;
     MM := TSQL.Detail[i].GetValue('MOIS');

     If MM < 10 Then Champ := 'M0'+IntToStr(MM) Else Champ := 'M'+IntToStr(MM);

     Tetab.Detail[y].PutValue(Champ,Tetab.Detail[y].GetValue(Champ) + 1);
     Ttot.Detail[y].PutValue(Champ,Ttot.Detail[y].GetValue(Champ) + 1);

     If (i = TSQL.Detail.Count -1) Or
        (TSQL.Detail[i].GetValue('PHC_ETABLISSEMENT')<>TSQL.Detail[i+1].GetValue('PHC_ETABLISSEMENT')) Then
     Begin
        For y := 0 To Tetab.Detail.Count -1 do
        Begin
           TFille := Tob.Create('Fille',TobEdit,-1);
           TFille.AddChampSupValeur('RUPTURE1',TSQL.Detail[i].GetValue('PHC_ETABLISSEMENT'));
           TFille.AddChampSupValeur('ETAB',TSQL.Detail[i].GetValue('PHC_ETABLISSEMENT')+' '+TSQL.Detail[i].GetValue('ET_LIBELLE'));
           TFille.AddChampSupValeur('TITRE',Tetab.Detail[y].GetValue('CC_LIBELLE'));
           TFille.AddChampSupValeur('NONMOUV',NonMouv);

           If GetCheckBoxState('PAIEDECALEE')= CbUnChecked then  //PT01
           Begin                                                 //PT01
              TFille.AddChampSupValeur('D01',Tetab.Detail[y].GetValue('M01'));
              TFille.AddChampSupValeur('D02',Tetab.Detail[y].GetValue('M02'));
              TFille.AddChampSupValeur('D03',Tetab.Detail[y].GetValue('M03'));
              TFille.AddChampSupValeur('D04',Tetab.Detail[y].GetValue('M04'));
              TFille.AddChampSupValeur('D05',Tetab.Detail[y].GetValue('M05'));
              TFille.AddChampSupValeur('D06',Tetab.Detail[y].GetValue('M06'));
              TFille.AddChampSupValeur('D07',Tetab.Detail[y].GetValue('M07'));
              TFille.AddChampSupValeur('D08',Tetab.Detail[y].GetValue('M08'));
              TFille.AddChampSupValeur('D09',Tetab.Detail[y].GetValue('M09'));
              TFille.AddChampSupValeur('D10',Tetab.Detail[y].GetValue('M10'));
              TFille.AddChampSupValeur('D11',Tetab.Detail[y].GetValue('M11'));
              TFille.AddChampSupValeur('D12',Tetab.Detail[y].GetValue('M12'));
           //PT01 Ajout ======>
           End
           Else
           Begin
              TFille.AddChampSupValeur('D01',Tetab.Detail[y].GetValue('M12'));
              TFille.AddChampSupValeur('D02',Tetab.Detail[y].GetValue('M01'));
              TFille.AddChampSupValeur('D03',Tetab.Detail[y].GetValue('M02'));
              TFille.AddChampSupValeur('D04',Tetab.Detail[y].GetValue('M03'));
              TFille.AddChampSupValeur('D05',Tetab.Detail[y].GetValue('M04'));
              TFille.AddChampSupValeur('D06',Tetab.Detail[y].GetValue('M05'));
              TFille.AddChampSupValeur('D07',Tetab.Detail[y].GetValue('M06'));
              TFille.AddChampSupValeur('D08',Tetab.Detail[y].GetValue('M07'));
              TFille.AddChampSupValeur('D09',Tetab.Detail[y].GetValue('M08'));
              TFille.AddChampSupValeur('D10',Tetab.Detail[y].GetValue('M09'));
              TFille.AddChampSupValeur('D11',Tetab.Detail[y].GetValue('M10'));
              TFille.AddChampSupValeur('D12',Tetab.Detail[y].GetValue('M11'));
           End;
           //PT01 Fin Ajout <=====
           Tetab.Detail[y].PutValue('M01',0);
           Tetab.Detail[y].PutValue('M02',0);
           Tetab.Detail[y].PutValue('M03',0);
           Tetab.Detail[y].PutValue('M04',0);
           Tetab.Detail[y].PutValue('M05',0);
           Tetab.Detail[y].PutValue('M06',0);
           Tetab.Detail[y].PutValue('M07',0);
           Tetab.Detail[y].PutValue('M08',0);
           Tetab.Detail[y].PutValue('M09',0);
           Tetab.Detail[y].PutValue('M10',0);
           Tetab.Detail[y].PutValue('M11',0);
           Tetab.Detail[y].PutValue('M12',0);
        End;
     End;
  End;
  For y := 0 To Ttot.Detail.Count -1 do
  Begin
     TFille := Tob.Create('Fille',TobEdit,-1);
     TFille.AddChampSupValeur('RUPTURE1','***');
     TFille.AddChampSupValeur('ETAB','TOTAL GENERAL');
     TFille.AddChampSupValeur('TITRE',Ttot.Detail[y].GetValue('CC_LIBELLE'));
     TFille.AddChampSupValeur('NONMOUV',NonMouv);

     If GetCheckBoxState('PAIEDECALEE')= CbUnChecked then  //PT01
     Begin                                                 //PT01
        TFille.AddChampSupValeur('D01',Ttot.Detail[y].GetValue('M01'));
        TFille.AddChampSupValeur('D02',Ttot.Detail[y].GetValue('M02'));
        TFille.AddChampSupValeur('D03',Ttot.Detail[y].GetValue('M03'));
        TFille.AddChampSupValeur('D04',Ttot.Detail[y].GetValue('M04'));
        TFille.AddChampSupValeur('D05',Ttot.Detail[y].GetValue('M05'));
        TFille.AddChampSupValeur('D06',Ttot.Detail[y].GetValue('M06'));
        TFille.AddChampSupValeur('D07',Ttot.Detail[y].GetValue('M07'));
        TFille.AddChampSupValeur('D08',Ttot.Detail[y].GetValue('M08'));
        TFille.AddChampSupValeur('D09',Ttot.Detail[y].GetValue('M09'));
        TFille.AddChampSupValeur('D10',Ttot.Detail[y].GetValue('M10'));
        TFille.AddChampSupValeur('D11',Ttot.Detail[y].GetValue('M11'));
        TFille.AddChampSupValeur('D12',Ttot.Detail[y].GetValue('M12'));
     //PT01 Ajout ======>
     End
     Else
     Begin
        TFille.AddChampSupValeur('D01',Ttot.Detail[y].GetValue('M12'));
        TFille.AddChampSupValeur('D02',Ttot.Detail[y].GetValue('M01'));
        TFille.AddChampSupValeur('D03',Ttot.Detail[y].GetValue('M02'));
        TFille.AddChampSupValeur('D04',Ttot.Detail[y].GetValue('M03'));
        TFille.AddChampSupValeur('D05',Ttot.Detail[y].GetValue('M04'));
        TFille.AddChampSupValeur('D06',Ttot.Detail[y].GetValue('M05'));
        TFille.AddChampSupValeur('D07',Ttot.Detail[y].GetValue('M06'));
        TFille.AddChampSupValeur('D08',Ttot.Detail[y].GetValue('M07'));
        TFille.AddChampSupValeur('D09',Ttot.Detail[y].GetValue('M08'));
        TFille.AddChampSupValeur('D10',Ttot.Detail[y].GetValue('M09'));
        TFille.AddChampSupValeur('D11',Ttot.Detail[y].GetValue('M10'));
        TFille.AddChampSupValeur('D12',Ttot.Detail[y].GetValue('M11'));
     End;
     //PT01 Fin Ajout <=====
  End;
  //TOBDebug(TSQL);
  //TOBDebug(Tetab);
  //TOBDebug(Ttot);
  //TOBDebug(TobEdit);
  TSQL.Free;
  Tetab.Free;
  Ttot.Free;
  TFQRS1(Ecran).LaTob:= TobEdit;
End;

//Evite le messsage : Impossible de Focaliser une fenetre désactivée ou invisible
// quand on quitte la fenetre par la croix alors qu'une date est en erreur
Procedure TOF_SUIVI_EFFECTIFS.OnClose ;
Var
  DateDefaut : THEdit;

Begin
  Inherited ;
  OnSort := True;
  DateDefaut := THEdit(GetControl('XX_VARIABLEDEB'));
  If DateDefaut <> nil Then
     If (not(IsValidDate(DateDefaut.text))) Then
        SetControltext('XX_VARIABLEDEB',DateToStr(Date));

  DateDefaut := THEdit(GetControl('XX_VARIABLEFIN'));
  If DateDefaut <> nil Then
     If (not(IsValidDate(DateDefaut.text))) Then
        SetControltext('XX_VARIABLEFIN',DateToStr(Date));

  FreeAndNil (TobEdit);
End ;

Procedure TOF_SUIVI_EFFECTIFS.Change(Sender: TObject);
Begin
  ControleCheckRupture;
End;

Procedure TOF_SUIVI_EFFECTIFS.ChangeExercice(Sender: TObject);
Begin
  TraitChangeExercice;
End;

Procedure TOF_SUIVI_EFFECTIFS.TraitChangeExercice ;
var
  SQL ,DebPer ,FinPer : String;
  Q   : TQuery;

Begin
  DebPer := DateToStr(idate1900);
  FinPer := DateToStr(idate1900);

  SQL := 'SELECT MIN(PPU_DATEDEBUT),MAX(PPU_DATEFIN) FROM PAIEENCOURS '+
         'LEFT JOIN EXERSOCIAL ON PPU_DATEDEBUT >= PEX_DATEDEBUT AND PPU_DATEFIN <= PEX_DATEFIN '+
         'WHERE PEX_EXERCICE = "'+GetControlText('EDEXERSOC')+'"';

  Q := OpenSQL(SQL,True);
  If not Q.eof Then
  Begin
     DebPer := Q.Fields[0].asstring;
     FinPer := Q.Fields[1].asstring;
  End;
  Ferme(Q);

  SetControlText('XX_VARIABLEDEB',DebPer);
  SetControlText('XX_VARIABLEFIN',FinPer);
End;

Procedure TOF_SUIVI_EFFECTIFS.VerifPeriode1(Sender: TObject) ;
var
  DebPer ,FinPer :THEdit;
  DatDeb ,DatFin :TDateTime;
  YYD,MMD,JJ,YYF,MMF : WORD;

Begin
  If OnSort = False Then
  Begin
    DebPer := THEdit(GetControl('XX_VARIABLEDEB'));
    FinPer := THEdit(GetControl('XX_VARIABLEFIN'));
    If (DebPer <> nil) And (FinPer <> nil) Then
    Begin
      DatDeb := StrToDate(DebPer.Text);
      DatFin := StrToDate(FinPer.Text);
      If DatDeb > DatFin Then
      Begin
        PGIBOX('La date saisie doit être inférieure à '+FinPer.Text+'',''+TFQRS1(Ecran).Caption+'');
        DebPer.SetFocus;
        Exit;
      End;
      DecodeDate(DatDeb,YYD,MMD,JJ);
      DecodeDate(DatFin,YYF,MMF,JJ);
      If (YYF>YYD) And (MMF>=MMD) Then
      Begin
        PgiBox('La période d''édition ne peut excéder douze mois civils.',''+TFQRS1(Ecran).Caption+'');
        DebPer.SetFocus;
      End;
    End;
  End;
End;

Procedure TOF_SUIVI_EFFECTIFS.VerifPeriode2(Sender: TObject) ;
var
  DebPer ,FinPer :THEdit;
  DatDeb ,DatFin :TDateTime;
  YYD,MMD,JJ,YYF,MMF : WORD;

Begin
  If OnSort = False Then
  Begin
    DebPer := THEdit(GetControl('XX_VARIABLEDEB'));
    FinPer := THEdit(GetControl('XX_VARIABLEFIN'));
    If (DebPer <> nil) And (FinPer <> nil) Then
    Begin
      DatDeb := StrToDate(DebPer.Text);
      DatFin := StrToDate(FinPer.Text);
      If DatDeb > DatFin Then
      Begin
        PGIBOX('La date saisie doit être supérieure à '+DebPer.Text+'',''+TFQRS1(Ecran).Caption+'');
        FinPer.SetFocus;
        Exit;
      End;
      DecodeDate(DatDeb,YYD,MMD,JJ);
      DecodeDate(DatFin,YYF,MMF,JJ);
      If (YYF>YYD) And (MMF>=MMD) Then
      Begin
        PgiBox('La période d''édition ne peut excéder douze mois civils.',''+TFQRS1(Ecran).Caption+'');
        DebPer.SetFocus;
      End;
    End;
  End;
End;

Procedure TOF_SUIVI_EFFECTIFS.ControleChampCompl ;
var
  I: Integer;
  Check: TCheckBox;

Begin
  If VH_Paie.PGLibCodeStat <> '' then
  Begin
     Check := TCheckBox(GetControl('CN5'));
     If Check <> nil Then
     Begin
       Check.Visible := True;
       Check.Enabled := True;
       Check.Caption := VH_Paie.PGLibCodeStat;
     End;
  End;
  For I := 1 to 4 do
  Begin
     If I <= VH_Paie.PGNbreStatOrg then
     Begin
       Check := TCheckBox(GetControl('CN'+IntToStr(I)));
       If Check <> nil Then
       Begin
         Check.Visible := True;
         Check.Enabled := True;
         If I = 1 then Check.Caption := VH_Paie.PGLibelleOrgStat1;
         If I = 2 then Check.Caption := VH_Paie.PGLibelleOrgStat2;
         If I = 3 then Check.Caption := VH_Paie.PGLibelleOrgStat3;
         If I = 4 then Check.Caption := VH_Paie.PGLibelleOrgStat4;
       End;
     End;
  End;
  For I := 1 to 4 do
  Begin
     If I <= VH_Paie.PgNbCombo then
     Begin
       Check := TCheckBox(GetControl('CL'+IntToStr(I)));
       If Check <> nil Then
       Begin
         Check.Visible := True;
         Check.Enabled := True;
         If I = 1 then Check.Caption := VH_Paie.PgLibCombo1;
         If I = 2 then Check.Caption := VH_Paie.PgLibCombo2;
         If I = 3 then Check.Caption := VH_Paie.PgLibCombo3;
         If I = 4 then Check.Caption := VH_Paie.PgLibCombo4;
       End;
     End;
  End;
End;

Procedure TOF_SUIVI_EFFECTIFS.ControleCheckRupture ;
var
  TabLieuTravail : array[1..10] of TCheckBox;
  PosCheck, PosUnCheck,i : integer;
  TH_Titre, TH_Champ : THEdit;
  Ok : boolean ;

Begin
  TabLieuTravail[1]:=TCheckBox(GetControl('CN1'));
  TabLieuTravail[2]:=TCheckBox(GetControl('CN2'));
  TabLieuTravail[3]:=TCheckBox(GetControl('CN3'));
  TabLieuTravail[4]:=TCheckBox(GetControl('CN4'));
  TabLieuTravail[5]:=TCheckBox(GetControl('CN5'));
  TabLieuTravail[6]:=TCheckBox(GetControl('CL1'));
  TabLieuTravail[7]:=TCheckBox(GetControl('CL2'));
  TabLieuTravail[8]:=TCheckBox(GetControl('CL3'));
  TabLieuTravail[9]:=TCheckBox(GetControl('CL4'));
  TabLieuTravail[10]:=nil;
  TH_Titre :=THEdit(GetControl('TITRE'));
  TH_Champ :=THEdit(GetControl('CHAMP'));
  PosUnCheck:=0;
  PosCheck:=0;

  For i:=1 to 9 do
     If (TabLieuTravail[i]<>nil) Then Ok:=False Else Begin Ok:=True; break; End;

  If Ok=False Then
  Begin
    //Coche une rupture
    For i:=1 to 9 do
       If (TabLieuTravail[i].checked=True) Then PosCheck:=i;
    If PosCheck > 0 Then
       For i:=1 to 9 do
          If i<>PosCheck then TabLieuTravail[i].enabled:=False;

    //Décoche une rupture ,  rend enable(True) les autres champs de rupture
    For i:=1 to 9 do
       If (TabLieuTravail[i].checked=False) and (TabLieuTravail[i].enabled=True) then PosUnCheck:=i;
    If (PosCheck=0) and (PosUnCheck>0) then
       For i:=1 to 9 do
          TabLieuTravail[i].enabled:=True;
  End;
  If (TH_Titre <> nil) And (TH_Champ <> nil) Then
  Begin
    TH_Titre.Text := '';
    TH_Champ.Text := '';
    CodeSel:= '';
    Case PosCheck Of
      1:Begin
         TH_Titre.Text := VH_Paie.PGLibelleOrgStat1;
         TH_Champ.Text := 'PHC_TRAVAILN1';
         CodeSel := 'PAG';
        End;
      2:Begin
         TH_Titre.Text := VH_Paie.PGLibelleOrgStat2;
         TH_Champ.Text := 'PHC_TRAVAILN2';
         CodeSel := 'PST';
        End;
      3:Begin
         TH_Titre.Text := VH_Paie.PGLibelleOrgStat3;
         TH_Champ.Text := 'PHC_TRAVAILN3';
         CodeSel := 'PUN';
        End;
      4:Begin
         TH_Titre.Text := VH_Paie.PGLibelleOrgStat4;
         TH_Champ.Text := 'PHC_TRAVAILN4';
         CodeSel := 'PSV';
        End;
      5:Begin
         TH_Titre.Text := VH_Paie.PGLibCodeStat;
         TH_Champ.Text := 'PHC_CODESTAT';
         CodeSel := 'PSQ';
        End;
      6:Begin
         TH_Titre.Text := VH_Paie.PgLibCombo1;
         TH_Champ.Text := 'PHC_LIBREPCMB1';
         CodeSel := 'PL1';
        End;
      7:Begin
         TH_Titre.Text := VH_Paie.PgLibCombo2;
         TH_Champ.Text := 'PHC_LIBREPCMB2';
         CodeSel := 'PL2';
        End;
      8:Begin
         TH_Titre.Text := VH_Paie.PgLibCombo3;
         TH_Champ.Text := 'PHC_LIBREPCMB3';
         CodeSel := 'PL3';
        End;
      9:Begin
         TH_Titre.Text := VH_Paie.PgLibCombo4;
         TH_Champ.Text := 'PHC_LIBREPCMB4';
         CodeSel := 'PL4';
        End;
    End;
  End;
End;

Initialization
  registerclasses ( [ TOF_SUIVI_EFFECTIFS ] ) ;
End.
