{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 25/05/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGHISTOGROUPEE ()
Mots clefs ... : TOF;PGHISTOGROUPEE
*****************************************************************
PT1 17/07/2007 JL V_800 Ajout saisie montant minimum pour augmentation en %
PT2 07/08/2007 FC V_800 Ajout de la saisie globale des éléments dynamiques et des éléments dossier
PT3 18/04/2008 GGU V_81 FQ 15361 Gestion uniformisée des zones libres - tables dynamiques
}
Unit UTofPGHistoGroupeeGen ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Graphics,
     HTB97,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     FE_Main,
     Fiche,
     EdtREtat,
{$else}
     eMul,
     eFiche,
     MainEAGL,
     UTilEAGL,
{$ENDIF}
     forms,
     uTob,
     sysutils,
     ComCtrls,
     ParamSoc,
     HCtrls,
     HEnt1,
     Vierge,
     HMsgBox,
     UTobDebug,
     HSysMenu,
     Grids,
     ed_tools,
     P5Util,
     ExtCtrls,
     Menus,
     ShellAPI,
     HPanel,
     windows,
     EntPaie,
     MailOl,
     UTofPGMulHistoGroupee,
     UTOFPGMULELTDYNGROUPEE, //PT2
     UTOF ;

Type
  TOF_PGHISTOGROUPEEGEN = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose  ; override ;
    private
    GHistorique : THGrid;
    procedure MiseEnFormeGrille;
    procedure ValiderSaisie (Sender : TObject);
  end ;

Type
  TOF_PGHISTOGROUPEENUM = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose  ; override ;
    private
    GHistorique : THGrid;
    procedure ChangeTypeCalc(Sender : TObject);
    procedure ValiderSaisie (Sender : TObject);
    Function  ArrondiPct(Montant : Double) : Double;
  end ;

Type
  TOF_PGELTDYNGROUPEEGEN = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose  ; override ;
    private
    GEltDyn : THGrid;
    DateApplic : TDateTime;
    procedure MiseEnFormeGrille;
    procedure ValiderSaisie (Sender : TObject);
  end ;

Implementation

Uses
  PGTablesDyna //PT3
  ;
procedure TOF_PGHISTOGROUPEEGEN.OnClose  ;
begin
     Inherited ;
end;

procedure TOF_PGHISTOGROUPEEGEN.OnArgument (S : String ) ;
var BVal : TToolBarButton97;
begin
  Inherited ;
  GHistorique := THGrid(GetControl('GMASSE'));
  MiseEnFormeGrille;
  BVal := TToolBarButton97(GetControl('BValider'));
  If BVal <> Nil then BVal.OnClick := ValiderSaisie;
  Ecran.Caption :=  'Saisie globale';
  UpdateCaption(Ecran);
end ;

procedure TOF_PGHISTOGROUPEEGEN.MiseEnFormeGrille;
var i,NumChamp : Integer;
    TypeChamp,Tablette,Libelle : String;
    Titres : HTStringList;
    HMTrad: THSystemMenu;
begin
     GHistorique.ColCount := PGTobLesChamps.Detail.Count;
     Titres := HTStringList.Create;
     For i := 0 to PGTobLesChamps.Detail.Count - 1 do
     begin
          NumChamp := PGTobLesChamps.Detail[i].GetValue('NUMERO');
          TypeChamp := PGTobLesChamps.Detail[i].GetValue('LETYPE');
          Tablette := PGTobLesChamps.Detail[i].GetValue('TABLETTE');
          Libelle := PGTobLesChamps.Detail[i].GetValue('LIBCHAMP');
          Titres.Insert(numchamp - 1, Libelle + 'Modifié');
     end;
     GHistorique.Titres := Titres;
     Titres.free;
     For i := 0 to PGTobLesChamps.Detail.Count - 1 do
     begin
          NumChamp := PGTobLesChamps.Detail[i].GetValue('NUMERO');
          TypeChamp := PGTobLesChamps.Detail[i].GetValue('LETYPE');
          Tablette := PGTobLesChamps.Detail[i].GetValue('TABLETTE');
          Libelle := PGTobLesChamps.Detail[i].GetValue('LIBCHAMP');
          if TypeChamp = 'D' then GHistorique.ColFormats[numchamp - 1] := ShortdateFormat
          else If TypeChamp = 'F' then
          begin
                GHistorique.ColFormats[numchamp - 1] := '# ##0.00';
                GHistorique.colTypes[numchamp - 1] := 'R';
                GHistorique.colAligns[numchamp - 1] := TaRightJustify;
          end
          else If TypeChamp = 'I' then
          begin
                GHistorique.ColFormats[numchamp - 1] := '# ##0';
                GHistorique.colTypes[numchamp - 1] := 'R';
                GHistorique.colAligns[numchamp - 1] := TaRightJustify;
          end
          else If TypeChamp = 'B' then
          begin
               GHistorique.colTypes[numchamp - 1] := 'B';
               GHistorique.ColFormats[numchamp - 1] := IntToStr (Ord (csCheckBox));
          end
          else If TypeChamp = 'T' then GHistorique.ColFormats[numchamp - 1] := 'CB='+Tablette;
     end;
     HMTrad.ResizeGridColumns(GHistorique) ;
end;

procedure TOF_PGHISTOGROUPEEGEN.ValiderSaisie (Sender : TObject);
var i,c : Integer;
    TypeChamp,Tablette,Champ,Indexchamp : String;
    ValApres : String;
begin
     NextPrevControl(TFFiche(Ecran));
     For i := 0 to PGTobHistoSal.Detail.Count - 1 do
     begin
          For c := 0 to PGTobLesChamps.Detail.Count - 1 do
          begin
               TypeChamp := PGTobLesChamps.Detail[c].GetValue('LETYPE');
               Tablette := PGTobLesChamps.Detail[c].GetValue('TABLETTE');
               Champ := PGTobLesChamps.Detail[c].GetValue('CHAMP');
               Indexchamp := PGTobLesChamps.Detail[c].GetValue('IDENT');
               ValApres := GHistorique.CellValues[c,1];
               If ValApres <> '' then PGTobHistoSal.detail[i].PutValue(Champ+'M',ValApres);
          end;
     end;
end;

{TOF_PGHISTOGROUPEENUM}

procedure TOF_PGHISTOGROUPEENUM.OnClose  ;
begin
     Inherited ;
end;

procedure TOF_PGHISTOGROUPEENUM.OnArgument (S : String ) ;
var Combo : THValComboBox;
    BVal : TToolBarButton97;
begin
  Inherited ;
  GHistorique := THGrid(GetControl('GMASSE'));
  BVal := TToolBarButton97(GetControl('BValider'));
  If BVal <> Nil then BVal.OnClick := ValiderSaisie;
  Ecran.Caption :=  'Saisie globale des montants';
  UpdateCaption(Ecran);
  Combo := THValComboBox(GetControl('TYPECALC'));
  if Combo <> Nil then Combo.OnChange := ChangeTypeCalc;
  SetControlText('TYPECALC','PCT');
  SetControlText('TYPEARRONDI','PROCHE');
end ;

procedure TOF_PGHISTOGROUPEENUM.ChangeTypeCalc(Sender : TObject);
var Titres : HTStringList;
    i,NbChamp : Integer;
    TypeChamp,Tablette,Libelle,Champ : String;
    HMTrad: THSystemMenu;
begin
     Titres := HTStringList.Create;
     Titres.Insert(0, 'Champ');
     Titres.Insert(1, 'Libellé');
     If GetControlText('TYPECALC') = 'PCT' then
     begin
      GHistorique.ColCount := 4;
      Titres.Insert(2, 'Pourcentage');
      Titres.Insert(3, 'Mt minimum'); //PT1
     end
     else
     begin
      GHistorique.ColCount := 3;
      Titres.Insert(2, 'Montant');
     end;
     GHistorique.Titres := Titres;
     Titres.free;
     SetControlEnabled('TYPEARRONDI',GetControlText('TYPECALC')='PCT');
     SetControlEnabled('NBDEC',GetControlText('TYPECALC')='PCT');
     NbChamp := 0;
     GHistorique.FixedCols := 2;
     GHistorique.ColWidths[0] := -1;
     GHistorique.ColWidths[1] := 150;
     GHistorique.ColWidths[2] := 80;
     GHistorique.ColFormats[2] := '# ##0.00';
     GHistorique.colTypes[2] := 'R';
     GHistorique.colAligns[2] := TaRightJustify;
     If GetControlText('TYPECALC') = 'PCT' then     //PT1
     begin
      GHistorique.ColWidths[3] := 80;
      GHistorique.ColFormats[3] := '# ##0.00';
      GHistorique.colTypes[3] := 'R';
      GHistorique.colAligns[3] := TaRightJustify;
     end;
     GHistorique.RowCount := PGTobLesChamps.Detail.Count + 1;
     For i := 0 to PGTobLesChamps.Detail.Count - 1 do
     begin
          TypeChamp := PGTobLesChamps.Detail[i].GetValue('LETYPE');
          Tablette := PGTobLesChamps.Detail[i].GetValue('TABLETTE');
          Libelle := PGTobLesChamps.Detail[i].GetValue('LIBCHAMP');
          Champ := PGTobLesChamps.Detail[i].GetValue('CHAMP');
          If (TypeChamp = 'F') or (TypeChamp = 'I') then
          begin
               NbChamp := NbChamp + 1;
               GHistorique.CellValues[0,NbChamp] := Champ;
               GHistorique.CellValues[1,NbChamp] := Libelle;
               GHistorique.CellValues[2,NbChamp] := '0';
               If GetControlText('TYPECALC') = 'PCT' then GHistorique.CellValues[3,NbChamp] := '0';    //PT1
          end;
     end;
     GHistorique.RowCount := NbChamp + 1;
     GHistorique.FixedRows := 1;
     HMTrad.ResizeGridColumns(GHistorique) ;
end;

procedure TOF_PGHISTOGROUPEENUM.ValiderSaisie (Sender : TObject);
var i,c : Integer;
    Champ : String;
    ValAvant,ValApres : String;
    Montant,ValAvantF,ValApresF,MontantMin,MontantAugm : Double;
begin
     NextPrevControl(TFFiche(Ecran));
     For i := 1 to GHistorique.RowCount - 1 do
     begin
          If IsNumeric (GHistorique.CellValues[2,i]) then Montant := StrToFloat(GHistorique.CellValues[2,i])
          Else Montant := 0;
          If GetControlText('TYPECALC') = 'PCT' then
          begin
           If IsNumeric (GHistorique.CellValues[3,i]) then MontantMin := StrToFloat(GHistorique.CellValues[3,i])
           Else MontantMin := 0;
          end;
          Champ := GHistorique.CellValues[0,i];
          If Montant <> 0 then
          begin
               For c := 0 to PGTobHistoSal.Detail.Count - 1 do
               begin
                    ValAvant := PGTobHistoSal.detail[c].GetValue(Champ);
                    if IsNumeric(ValAvant) then ValAvantF := StrToFloat(ValAvant)
                    Else ValAvantF := 0;
                    If GetControlText('TYPECALC') = 'PCT' then
                    begin
                    //DEBUT PT1
                      MontantAugm := (ValAVantF*Montant)/100;
                      If MontantMin > 0 then
                      begin
                        If MontantAugm < MontantMin then MontantAugm := MontantMin;
                      end;
                      ValApresF := ArrondiPct(ValAVantF + MontantAugm);
                    //FIN PT1
                    end
                    else ValApresF := ValAVantF + Montant;
                    ValApres := FloatToStr(ValApresF);
                    PGTobHistoSal.detail[c].PutValue(Champ+'M',ValApres);
               end;
          end;
     end;
end;

Function TOF_PGHISTOGROUPEENUM.ArrondiPct(Montant : Double) : Double;
var CalculEntier,Calcul,MontantDec : Double;
    NbDec : Integer;
    TypeArrondi : String;
    SpDec : THSpinEdit;
begin
    SpDec := THSpinEdit(GetControl('NBDEC'));
    If SpDec <> Nil then NbDec := SpDec.Value
    else NbDec := 2;
    MontantDec := Montant;
    If NbDec = 1 then MontantDec := Montant * 10
    else If NbDec = 2 then MontantDec := Montant * 100;
    If THEdit(GetControl('TYPEARRONDI')) <> Nil then TypeArrondi := GetControlText('TYPEARRONDI')
    else TypeArrondi := 'PROCHE';
    If TypeArrondi = 'PROCHE' then Calcul := (arrondi(MontantDec,0))
    else If TypeArrondi = 'SUP' then
    begin
        CalculEntier := Int(MontantDec);
        If MontantDec > CalculEntier then Calcul := CalculEntier + 1
        else Calcul := CalculEntier
    end
    else If TypeArrondi = 'INF' then Calcul := Int(MontantDec);
    If NbDec = 1 then Calcul := Calcul / 10
    else If NbDec = 2 then Calcul := Calcul / 100;
    Result := Calcul;
end;

//DEB PT2
{ TOF_PGELTDYNGROUPEEGEN }
procedure TOF_PGELTDYNGROUPEEGEN.OnClose  ;
begin
  Inherited ;
end;

procedure TOF_PGELTDYNGROUPEEGEN.OnArgument (S : String ) ;
var BVal : TToolBarButton97;
begin
  Inherited ;
  DateApplic := StrToDateTime(S);
  GEltDyn := THGrid(GetControl('GMASSE'));

  MiseEnFormeGrille;

  BVal := TToolBarButton97(GetControl('BValider'));
  If BVal <> Nil then BVal.OnClick := ValiderSaisie;
  if (PGTypElement = 'ELD') then
    Ecran.Caption :=  'Saisie globale des éléments dynamiques'
  else
    Ecran.Caption :=  'Saisie globale des éléments dossier';
  UpdateCaption(Ecran);
end ;

procedure TOF_PGELTDYNGROUPEEGEN.MiseEnFormeGrille;
var i : Integer;
    PgTypeDonne,CodTabl,Libelle : String;
    Titres : HTStringList;
    HMTrad: THSystemMenu;
begin
     GEltDyn.ColCount := PGTobLesElements.Detail.Count;
     Titres := HTStringList.Create;
     For i := 0 to PGTobLesElements.Detail.Count - 1 do
     begin
       Libelle := PGTobLesElements.Detail[i].GetValue('LIBELLE');
       Titres.Insert(i, Libelle);
     end;
     GEltDyn.Titres := Titres;
     Titres.free;
     For i := 0 to PGTobLesElements.Detail.Count - 1 do
     begin
       if (PGTypElement = 'ELD') then
       begin
         PgTypeDonne := PGTobLesElements.Detail[i].GetValue('PGTYPEDONNE');
         CodTabl := PGTobLesElements.Detail[i].GetValue('CODTABL');
       end
       else
         PgTypeDonne := 'F';

       if PgTypeDonne = 'D' then
         GEltDyn.ColFormats[i] := ShortdateFormat
       else If PgTypeDonne = 'B' then
       begin
         GEltDyn.ColTypes[i] := 'B';
         GEltDyn.ColFormats[i] := IntToStr (Ord (csCheckBox));
         GEltDyn.ColAligns[i] := taCenter;
       end
       else If PgTypeDonne = 'T' then //PT3
{//@@GGU A étudier pour la FQ 15361 : Chaque salarié n'a pas forcément le même filtre,
il faudrait recalculer le plus à chaque changement de ligne, c'est à dire à dans le
onCellEnter
         GEltDyn.ColFormats[i] := 'CB=PGCOMBOZONELIBRE|'
                                + GetPlusPGCOMBOZONELIBRE(DateApplic, CodTabl, SalarieDeLaLigne) //PT3
}
         GEltDyn.ColFormats[i] := 'CB=PGCOMBOZONELIBRE|PTD_CODTABL="' + CodTabl +'" AND PTD_DTVALID=(SELECT MAX(PTD_DTVALID) FROM TABLEDIMDET WHERE PTD_CODTABL="' + CodTabl + '" AND PTD_DTVALID<="' + USDATETIME(DateApplic) + '")'
       else If PgTypeDonne = 'F' then
       begin
         GEltDyn.ColFormats[i] := '# ##0.000';
         GEltDyn.ColTypes[i] := 'R';
         GEltDyn.ColAligns[i] := TaRightJustify;
       end;
     end;
     HMTrad.ResizeGridColumns(GEltDyn) ;
end;

procedure TOF_PGELTDYNGROUPEEGEN.ValiderSaisie (Sender : TObject);
var i,c : Integer;
    PgInfosModif, CodeElt : String;
    ValApres : String;
begin
  NextPrevControl(TFFiche(Ecran));
  For i := 0 to PGTobEltDyn.Detail.Count - 1 do
  begin
    For c := 0 to PGTobLesElements.Detail.Count - 1 do
    begin
      if (PGTypElement = 'ELD') then
        PgInfosModif := PGTobLesElements.Detail[c].GetValue('PGINFOSMODIF')
      else
        CodeElt := PGTobLesElements.Detail[c].GetValue('CODEELT');
      ValApres := GEltDyn.CellValues[c,1];
      If ValApres <> '' then
      begin
        if (PGTypElement = 'ELD') then
          PGTobEltDyn.detail[i].PutValue(PgInfosModif,ValApres)
        else
          PGTobEltDyn.detail[i].PutValue(CodeElt,ValApres);
      end;
    end;
  end;
end;
//FIN PT2

Initialization
  registerclasses ( [ TOF_PGHISTOGROUPEEGEN,TOF_PGHISTOGROUPEENUM, TOF_PGELTDYNGROUPEEGEN ] ) ;
end.

