{***********UNITE*************************************************
Auteur  ...... : FC
Créé le ...... : 31/07/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGELTDYNGROUPEE ()
Mots clefs ... : TOF;PGELTDYNGROUPEE
*****************************************************************
PT1  04/12/2007 FC V_81 FQ  Pb filtrage des valeurs d'un élément quand plusieurs prédéfinis paramétrés
PT2  21/01/2008 FC V_81 FQ 15144 Elément dynamique => Typeinfols=ZLS au lieu de SAL
PT3  17/04/2008 GGU V81 FQ 15361 Factorisation du code en vue de la modification
}
Unit UTOFPGELTDYNGROUPEE ;

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
     UTobXls,
     UTofPGMulEltDynGroupee,
     UTOM,
     ParamDat,
     PGOutils2,
     PGOutilsHistorique,
     UTOF,
     LookUp;
Const
     ColSal = 0;
     ColNom = 1;
     ColDateV = 2;

Type
  TOF_PGELTDYNGROUPEE = Class (TOF)
  public
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose  ; override ;
  private
    GEltDyn : THGrid;
    NbElements : Integer;
    StChampGrid : String;
    TitreColonne: TStringList;
    procedure MiseEnFormeGrille;
    procedure RemplirGrille;
    procedure ZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
    function ZoneAccessible(var ACol, ARow: Longint): Boolean;
    procedure GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GrilleCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
    procedure ValiderSaisie (Sender : TObject);
    procedure SaisieGlobale(Sender : TObject);
    procedure GrilleCopierColler(Fichier :String);
    procedure BFichierClick(Sender : TObject);
    procedure BLegendeClick(Sender : TObject);
    procedure KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GrilleElipsisClick(Sender: TObject);
    procedure GrilleColEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    Function FormatageDouble(Chaine : String) : Double;
    Procedure changementDate(Sender : TObject);
    procedure GrilleKeyPress(Sender : TObject; var Key : Char);
//    procedure ImprimeListe(Sender : TObject);
  end ;

Implementation

uses TntGrids,
  PGTablesDyna; //PT3


Function MyStrToFloat(Str: String) : Extended;
var
  FormatSet : TFormatSettings;
begin
  GetLocaleFormatSettings(0,FormatSet);
  result := 0;
  if (Str <> '') and (IsNumeric(Str)) then
  begin
    try
      FormatSet.DecimalSeparator := '.';
      result := StrToFloat(Str,FormatSet);
    except
      FormatSet.DecimalSeparator := ',';
      result := StrToFloat(Str,FormatSet);
    end;
  end;
end;

procedure TOF_PGELTDYNGROUPEE.OnClose  ;
begin
  Inherited ;
  if TitreColonne <> nil then TitreColonne.Free;
end;

procedure TOF_PGELTDYNGROUPEE.OnArgument (S : String ) ;
var
  BLegende,BFichier,BVal,BSaisie, Bt{, BImprimeListe} : TToolBarButton97;
  NumChamp,i, ColAp : Integer;
  PgTypeDonne : String;
begin
  Inherited ;
  if (PGTypElement = 'ELD') then
    Ecran.Caption := 'Saisie groupée des éléments dynamiques'
  else
    Ecran.Caption := 'Saisie groupée des éléments dossier';
  UpdateCaption(Ecran);

  NumChamp := 1;
  StChampGrid := 'SALARIE;LIBELLE;DATEAPPLIC';
  NbElements := PGTobLesElements.detail.Count;
  For i := 0 to NbElements - 1 do
  begin
    if (PGTypElement = 'ELD') then
      StChampGrid := StChampGrid+';'+PGTobLesElements.detail[i].GetValue('PGINFOSMODIF')
    else
      StChampGrid := StChampGrid+';'+PGTobLesElements.detail[i].GetValue('CODEELT');
    StChampGrid := StChampGrid+';NUMCHAMP'+IntToStr(NumChamp);
    NumChamp := NumChamp + 1;
  end;
  GEltDyn := THGrid(GetControl('GELTDYN'));
  TitreColonne := TStringList.Create;

  GEltDyn.OnElipsisClick := GrilleElipsisClick;
  GEltDyn.OnColEnter := GrilleColEnter;
  GEltDyn.OnCellEnter := GrilleCellEnter;
  GEltDyn.OnCellExit := GrilleCellExit;
  GEltDyn.OnKeyPress := GrilleKeyPress;

//  BImprimeListe := TToolBarButton97(GetControl('BImprimeListe'));
//  If BImprimeListe <> Nil then BImprimeListe.OnClick := ImprimeListe;
  BVal := TToolBarButton97(GetControl('BVALIDSAISIE'));
  If BVal <> Nil then BVal.OnClick := ValiderSaisie;
  BSaisie := TToolBarButton97(GetControl('BSAISIE'));
  If BSaisie <> Nil then BSaisie.OnClick := SaisieGlobale;
  BFichier := TToolBarButton97(GetControl('BFICHIER'));
  If BFichier <> Nil then BFichier.OnClick := BFichierClick;
  GEltDyn.OnKeyDown := KeyDown;
  BLegende := TToolBarButton97(GetControl('BLEGENDE'));
  If BLegende <> Nil then BLegende.OnClick := BLegendeClick;
  SetControlText('DATEEFFET',DateToStr(DebutDeMois(V_PGI.DateEntree)));
  Bt :=  TToolBarButton97(GetControl('BDATE'));
  If Bt <> Nil then Bt.OnClick := ChangementDate;
  
  MiseEnFormeGrille;
  RemplirGrille;

  For i := 1 to GEltDyn.RowCount - 1 do
  begin
    GEltDyn.Cellvalues[2,i] := GetControlText('DATEEFFET');
  end;
  PGTobEltDyn.GetGridDetail(GEltDyn,GEltDyn.RowCount-1,'',StChampGrid);

  For i := 0 to PGTobLesElements.Detail.Count - 1 do
  begin
    NumChamp := PGTobLesElements.Detail[i].GetValue('NUMCHAMP');
    ColAp := 1 + (NumChamp * 2);
    if (PGTypElement = 'ELD') then
    begin
      PgTypeDonne := PGTobLesElements.Detail[i].GetValue('PGTYPEDONNE');
      if PgTypeDonne = 'D' then
        GEltDyn.ColFormats[ColAp] := ShortdateFormat;
{      else If PgTypeDonne = 'B' then
      begin
        GEltDyn.colTypes[ColAP] := 'B';
        GEltDyn.ColFormats[ColAp] := IntToStr(Ord(csCheckBox));
        GEltDyn.ColAligns[ColAP] := taCenter;
      end;}
    end
    else
      GEltDyn.ColFormats[ColAp] := '# ##0.000';
  end;
end ;

procedure TOF_PGELTDYNGROUPEE.MiseEnFormeGrille;
var i,NumChamp : Integer;
    LeElement,CodTabl,Libelle, PgInfosModif, PgTypeDonne : String;
    Titres : HTStringList;
    ColAp,ColOrdre,NbInfos : Integer;
begin
     TitreColonne.Clear;
     GEltDyn.ColCount := 3 + (PGTobLesElements.Detail.Count*2);
     Titres := HTStringList.Create;
     GEltDyn.ColEditables[ColSal] := False;
     GEltDyn.ColFormats[ColSal] := '';
     Titres.Insert(ColSal, 'Salarie');
     TitreColonne.Add('SALARIE');
     GEltDyn.ColEditables[ColNom] := False;
     GEltDyn.ColFormats[ColNom] := '';
     Titres.Insert(ColNom, 'Nom');
     TitreColonne.Add('LIBELLE');
     GEltDyn.ColWidths[ColNom] := 120;
     Titres.Insert(ColDateV, 'Date');
     TitreColonne.Add('DATEVAL');
     GEltDyn.ColWidths[ColDateV] := 100;
     GEltDyn.ColFormats[ColDateV] := ShortDateFormat;
     PGTobLesElements.Detail.Sort('NUMCHAMP');
     NbInfos := PGTobLesElements.Detail.Count;
     For i := 0 to PGTobLesElements.Detail.Count - 1 do
     begin
         NumChamp := PGTobLesElements.Detail[i].GetValue('NUMCHAMP');
         if (PGTypElement = 'ELD') then
         begin
           ColAp := 1 + (NumChamp * 2);
           ColOrdre := 2 + (NumChamp * 2);
           PgTypeDonne := PGTobLesElements.Detail[i].GetValue('PGTYPEDONNE');
           CodTabl := PGTobLesElements.Detail[i].GetValue('CODTABL');
           Libelle := PGTobLesElements.Detail[i].GetValue('LIBELLE');
           PgInfosModif := PGTobLesElements.Detail[i].GetValue('PGINFOSMODIF');
         end
         else
         begin
           ColAp := 1 + (NumChamp * 2);
           ColOrdre := 2 + (NumChamp * 2);
           PgTypeDonne := 'F';
           Libelle := PGTobLesElements.Detail[i].GetValue('LIBELLE');
         end;
         GEltDyn.ColWidths[ColOrdre] := -1;
         If NbInfos < 3 then
           GEltDyn.ColWidths[ColAp] := 100
         else
           GEltDyn.ColWidths[ColAp] := 90;
         Titres.Insert(ColAp, Libelle);
         TitreColonne.Add(LeElement);
         Titres.Insert(ColOrdre, 'NumChamp');
         TitreColonne.Add('NUMCHAMP'+IntToStr(NumChamp));
         if PgTypeDonne = 'D' then
         begin
           GEltDyn.ColFormats[ColAp] := ShortdateFormat;
           GEltDyn.ColTypes[ColAp] := 'D';
         end
{         else If PgTypeDonne = 'B' then
         begin
           GEltDyn.ColTypes[ColAP] := 'B';
           GEltDyn.ColFormats[ColAP] := IntToStr(Ord(csCheckBox));
         end}
//PT1         else if PgTypeDonne = 'T' then
//PT1           GEltDyn.ColFormats[ColAp] := 'CB=PGCOMBOZONELIBRE|PTD_CODTABL="' + CodTabl +'" AND PTD_DTVALID=(SELECT MAX(PTD_DTVALID) FROM TABLEDIMDET WHERE PTD_CODTABL="' + CodTabl + '" AND PTD_DTVALID<="' + USDATETIME(StrToDateTime(GetControlText('DATEEFFET'))) + '")'
         else If PgTypeDonne = 'F' then
           GEltDyn.ColFormats[ColAp] := '# ##0.000';
     end;
     GEltDyn.Titres := Titres;
     Titres.free;
end;

procedure TOF_PGELTDYNGROUPEE.GrilleElipsisClick(Sender: TObject);
var
  key : char;
  T : Tob;
  LeType : String;
  NumChamp : Double;
//PT6  St,
  StWhere:String;
//PT6  Q:TQuery;
//PT6  ConvSal,Etab : String;
  i : integer;
//PT6  DateMax:TDateTime;
begin
  if (PGTypElement = 'ELD') then
  begin
    //DEB PT1
    if GEltDyn.Col = 2 then
    begin
      key := '*';
      ParamDate (Ecran, Sender, Key);
    end
    //FIN PT1
    else
    begin
      NumChamp := (GEltDyn.Col - 1)/2;
      T := PGTobLesElements.FindFirst(['NUMCHAMP'],[NumChamp],False);
      if T <> Nil then
      begin
        LeType := T.GetValue('PGTYPEDONNE');
        if (LeType = 'D') then
        begin
          key := '*';
          ParamDate (Ecran, Sender, Key);
        end;
        //DEB PT1
        if (LeType = 'T') then
        begin
          i := GEltDyn.row; 
          StWhere := ' SELECT PTD_VALCRIT1,PTD_LIBELLECODE,PTD_RESULTAT FROM TABLEDIMDET '
                   + GetPlusPGCOMBOZONELIBRE(StrToDateTime(GEltDyn.CellValues[2,i]), T.GetValue('CODTABL'), GEltDyn.CellValues[0,i]); //PT3

{//PT3
          Q := OpenSQL('SELECT PSA_CONVENTION,PSA_ETABLISSEMENT FROM SALARIES WHERE PSA_SALARIE = "' + GEltDyn.CellValues[0,i] + '"',True,1);
          if not Q.Eof then
          begin
            ConvSal := Q.FindField('PSA_CONVENTION').AsString;
            Etab := Q.FindField('PSA_ETABLISSEMENT').AsString;
          end;
          Ferme(Q);

          //Vérifier s'il n'existe pas des valeurs pour l'établissement du salarié
          if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' +
          ' WHERE PTE_PREDEFINI="DOS" AND PTE_NODOSSIER="' + V_PGI.NoDossier + '"' +
          ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GEltDyn.CellValues[2,i]))+'"' +
          ' AND PTE_CODTABL="'+T.GetValue('CODTABL')+'"' +
          ' AND PTE_NIVSAIS="ETB" AND PTE_VALNIV="' + Etab + '"') then
          begin
            St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT WHERE PTE_CODTABL="'+T.GetValue('CODTABL')+'" '+
              ' AND PTE_PREDEFINI="DOS" AND PTE_NODOSSIER="' + V_PGI.NoDossier + '"' +
              ' AND PTE_NIVSAIS="ETB" AND PTE_VALNIV="' + Etab + '"' +
              ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GEltDyn.CellValues[2,i])) + '"';
            Q := OpenSQL(St,True,1);
            DateMax := iDate1900;
            if not Q.Eof then
              DateMax := Q.FindField('DTVALID').AsDateTime;
            Ferme(Q);
            StWhere := ' SELECT PTD_VALCRIT1,PTD_LIBELLECODE,PTD_RESULTAT FROM TABLEDIMDET ' +
              ' WHERE PTD_DTVALID = "' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+T.GetValue('CODTABL')+'"' +
              ' AND PTD_PREDEFINI="DOS" AND PTD_NODOSSIER="' + V_PGI.NoDossier + '"' +
              ' AND PTD_NIVSAIS="ETB" AND PTD_VALNIV="' + Etab + '"';
          end
          //Vérifier s'il n'existe pas des valeurs pour le dossier en général
          else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' +
          ' WHERE PTE_PREDEFINI="DOS" AND PTE_NODOSSIER="' + V_PGI.NoDossier + '"' +
          ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GEltDyn.CellValues[2,i]))+'"' +
          ' AND PTE_CODTABL="'+T.GetValue('CODTABL')+'"' +
          ' AND PTE_NIVSAIS="GEN"') then
          begin
            St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT WHERE PTE_CODTABL="'+T.GetValue('CODTABL')+'" '+
              ' AND PTE_PREDEFINI="DOS" AND PTE_NODOSSIER="' + V_PGI.NoDossier + '"' +
              ' AND PTE_NIVSAIS="GEN"' +
              ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GEltDyn.CellValues[2,i]))+'"';
            Q := OpenSQL(St,True,1);
            DateMax := iDate1900;
            if not Q.Eof then
              DateMax := Q.FindField('DTVALID').AsDateTime;
            Ferme(Q);
            StWhere := ' SELECT PTD_VALCRIT1,PTD_LIBELLECODE,PTD_RESULTAT FROM TABLEDIMDET ' +
              ' WHERE PTD_DTVALID = "' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+T.GetValue('CODTABL')+'"' +
              ' AND PTD_PREDEFINI="DOS" AND PTD_NODOSSIER="' + V_PGI.NoDossier + '"' +
              ' AND PTD_NIVSAIS="GEN"';
          end
          //Vérifier s'il n'existe pas des valeurs pour STD + convention
          else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' +
          ' WHERE PTE_PREDEFINI="STD"' +
          ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GEltDyn.CellValues[2,i]))+'"' +
          ' AND PTE_CODTABL="'+T.GetValue('CODTABL')+'"' +
          ' AND PTE_NIVSAIS="CON" AND PTE_VALNIV="' + ConvSal + '"') then
          begin
            St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT WHERE PTE_CODTABL="'+T.GetValue('CODTABL')+'" '+
              ' AND PTE_PREDEFINI="STD"' +
              ' AND PTE_NIVSAIS="CON" AND PTE_VALNIV="' + ConvSal + '"' +
              ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GEltDyn.CellValues[2,i]))+'"';
            Q := OpenSQL(St,True,1);
            DateMax := iDate1900;
            if not Q.Eof then
              DateMax := Q.FindField('DTVALID').AsDateTime;
            Ferme(Q);
            StWhere := ' SELECT PTD_VALCRIT1,PTD_LIBELLECODE,PTD_RESULTAT FROM TABLEDIMDET ' +
              ' WHERE PTD_DTVALID = "' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+T.GetValue('CODTABL')+'"' +
              ' AND PTD_PREDEFINI="STD"' +
              ' AND PTD_NIVSAIS="CON" AND PTD_VALNIV="' + ConvSal + '"';
          end
          //Vérifier s'il n'existe pas des valeurs pour STD + Convention 000
          else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' +
          ' WHERE PTE_PREDEFINI="STD"' +
          ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GEltDyn.CellValues[2,i]))+'"' +
          ' AND PTE_CODTABL="'+T.GetValue('CODTABL')+'"' +
          ' AND PTE_NIVSAIS="CON" AND PTE_VALNIV="000"') then
          begin
            St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT WHERE PTE_CODTABL="'+T.GetValue('CODTABL')+'" '+
              ' AND PTE_PREDEFINI="STD"' +
              ' AND PTE_NIVSAIS="CON" AND PTE_VALNIV="000"' +
              ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GEltDyn.CellValues[2,i]))+'"';
            Q := OpenSQL(St,True,1);
            DateMax := iDate1900;
            if not Q.Eof then
              DateMax := Q.FindField('DTVALID').AsDateTime;
            Ferme(Q);
            StWhere := ' SELECT PTD_VALCRIT1,PTD_LIBELLECODE,PTD_RESULTAT FROM TABLEDIMDET ' +
              ' WHERE PTD_DTVALID = "' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+T.GetValue('CODTABL')+'"' +
              ' AND PTD_PREDEFINI="STD"' +
              ' AND PTD_NIVSAIS="CON" AND PTD_VALNIV="000"';
          end
          //Vérifier s'il n'existe pas des valeurs pour STD + GEN
          else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' +
          ' WHERE PTE_PREDEFINI="STD"' +
          ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GEltDyn.CellValues[2,i]))+'"' +
          ' AND PTE_CODTABL="'+T.GetValue('CODTABL')+'"' +
          ' AND PTE_NIVSAIS="GEN"') then
          begin
            St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT WHERE PTE_CODTABL="'+T.GetValue('CODTABL')+'" '+
              ' AND PTE_PREDEFINI="STD" AND PTE_NIVSAIS="GEN"' +
              ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GEltDyn.CellValues[2,i]))+'"';
            Q := OpenSQL(St,True,1);
            DateMax := iDate1900;
            if not Q.Eof then
              DateMax := Q.FindField('DTVALID').AsDateTime;
            Ferme(Q);
            StWhere := ' SELECT PTD_VALCRIT1,PTD_LIBELLECODE,PTD_RESULTAT FROM TABLEDIMDET ' +
              ' WHERE PTD_DTVALID = "' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+T.GetValue('CODTABL')+'"' +
              ' AND PTD_PREDEFINI="STD" AND PTD_NIVSAIS="GEN"';
          end
          //Vérifier s'il n'existe pas des valeurs pour CEG
          else
          begin
            St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT WHERE PTE_CODTABL="'+T.GetValue('CODTABL')+'" '+
              ' AND PTE_PREDEFINI="CEG" AND PTE_NIVSAIS="GEN"' +
              ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GEltDyn.CellValues[2,i]))+'"';
            Q := OpenSQL(St,True,1);
            DateMax := iDate1900;
            if not Q.Eof then
              DateMax := Q.FindField('DTVALID').AsDateTime;
            Ferme(Q);
            StWhere := ' SELECT PTD_VALCRIT1,PTD_LIBELLECODE,PTD_RESULTAT FROM TABLEDIMDET ' +
              ' WHERE PTD_DTVALID = "' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+T.GetValue('CODTABL')+'"' +
              ' AND PTD_PREDEFINI="CEG" AND PTD_NIVSAIS="GEN"';
          end;
}
          LookUpList (GEltDyn,'Eléments dynamiques','','PTD_VALCRIT1','','','PTD_LIBELLECODE',TRUE,-1,StWhere);
        end;
        //FIN PT1
      end;
    end;
  end;
end;

procedure TOF_PGELTDYNGROUPEE.RemplirGrille;
begin
  PGTobEltDyn.PutGridDetail(GEltDyn,False,False,StChampGrid,False);
  GEltDyn.RowCount := PGTobEltDyn.Detail.Count + 1;
end;

procedure TOF_PGELTDYNGROUPEE.ZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
var
  Sens, i: integer;
  OldEna: boolean;
begin
  OldEna := GEltDyn.SynEnabled;
  GEltDyn.SynEnabled := False;
  Sens := -1;
  if GEltDyn.Row > ARow then
    Sens := 1
  else if ((GEltDyn.Row = ARow) and (ACol < GEltDyn.Col)) then
    Sens := 1;
  if (sens = -1) and ((ACol = ColSal) or (GEltDyn.Col = ColSal)) and (ARow = 1) then
  begin
    GEltDyn.SynEnabled := OldEna;
    Cancel := TRUE;
    exit;
  end;
  ACol := GEltDyn.Col;
  ARow := GEltDyn.Row;
  i := 0;
  while not ZoneAccessible(ACol, ARow) do
  begin
    Cancel := True;
    inc(i);
    if i > 1000 then Break;
    if (ACol = ColSal) and (ARow = 1) then
    begin
      ACol := 1;
      break;
    end;
    if Sens = 1 then
    begin
      if ((ACol = GEltDyn.ColCount - 1) and (ARow = GEltDyn.RowCount - 1)) then
      begin
        ACol := GEltDyn.FixedCols;
        ARow := 1;
        Break;
      end;
      if ACol < GEltDyn.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        ACol := GEltDyn.FixedCols;
      end;
    end else
    begin
      if ((ACol = GEltDyn.FixedCols) and (ARow = 1)) then Break;
      if ACol > GEltDyn.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := GEltDyn.ColCount - 1;
      end;
    end;
  end;
  GEltDyn.SynEnabled := OldEna;
end;

function TOF_PGELTDYNGROUPEE.ZoneAccessible(var ACol, ARow: Longint): Boolean;
var
  T1: TOB;
  i : integer;
begin
  result := FALSE;
  If (ACol = ColNom) or (ACol = ColDateV) then
  begin
       Result := True;
       Exit;
  end;
  For i := 1 to NbElements do
  begin
     If ACol = (1 + (i * 2))  then
     begin
        Result := True;
        Exit;
     end;
  end;
  If ACol = ColNom then
  begin
       Result := True;
       Exit;
  end;
  T1 := TOB(GEltDyn.Objects[ColSal, GEltDyn.Row]);
  if T1 = nil then
  begin
    result := FALSE;
    exit;
  end;
end;

procedure TOF_PGELTDYNGROUPEE.GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  ZoneSuivanteouOk(ACol, ARow, Cancel);
end;

procedure TOF_PGELTDYNGROUPEE.GrilleCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var
  TypeElement,Valeur : String;
  T : Tob;
  NumChamp : Double;
begin
  NumChamp := (ACol - 1)/2;
  T := PGTobLesElements.FindFirst(['NUMCHAMP'],[NumChamp],False);
  if (PGTypElement = 'ELD') then
  begin
    TypeElement := '';
    If T <> Nil then
      TypeElement := T.GetValue('PGTYPEDONNE');

    If TypeElement = 'D' then
    begin
      Valeur := GEltDyn.CellValues[ACol,ARow];
      If (Valeur = '__/__/____') or (Valeur = '  /  /    ') then
      begin
        Valeur := '';
        GEltDyn.CellValues[ACol,ARow] := '';
      end;
      If Valeur <> '' then
      begin
        If Not iSValidDate(Valeur) then
        begin
          PGIBox(Valeur + ' n''est pas une date valide');
          GEltDyn.CellValues[ACol,ARow] := '';
        end;
      end;
    end;
  end
  else
  begin
    if (ACol > 2) then
    begin
      Valeur := GEltDyn.CellValues[ACol,ARow];
      If Valeur <> '' then
      begin
        If Not IsNumeric(Valeur) then
        begin
          PGIBox(Valeur + ' n''est pas une valeur numérique correcte');
          GEltDyn.CellValues[ACol,ARow] := '';
        end;
      end;
    end;
  end;
end;

procedure TOF_PGELTDYNGROUPEE.ValiderSaisie (Sender : TObject);
var
  i,c : Integer;
  TobHisto,TH : Tob;
  Salarie,Etab,Popul,ConvSal : String;
  DateAp,DateFinVal : TDateTime;
  PgTypeDonne, CodTabl, PgInfosModif, CodeElt : String;
  ValApres : String;
  ColAp,NumChamp : Integer;
  TestDate : TDateTime;
  Req,St : String;
  Q,Qt : TQuery;
  LocalTob,TFind :Tob;
  ValComp : Double;
  BTrouve,BStop : Boolean;
  DateValidite : TDateTime;
  TypeNiveau : String;
  ACol,ARow:integer;
  tmp : boolean;
begin
  TraceErr:= TListBox (GetControl ('LSTBXERROR'));
  if TraceErr = NIL then
  begin
    PgiBox ('Attention, composant trace non trouvé, abandon du traitement',Ecran.Caption);
    exit;
  end;
  TraceErr.Refresh ;
  TraceErr.Clear;

  ACol := 2;
  if ACol = GEltDyn.Col then
    ACol := 3;
  GEltDyn.ForceCellEnter(ACol,1);

//  ACol := GEltDyn.Col;
//  ARow := GEltDyn.Row;
//  tmp := False;
//  GrilleCellExit(GEltDyn,ACol,ARow,tmp);

  NextPrevControl(TFFiche(Ecran));

  SetFocusControl('DATEEFFET');

  TestDate := V_PGI.DateEntree;
  InitMoveProgressForm (NIL,'Sauvegarde en cours',
                  'Veuillez patienter SVP ...', GEltDyn.RowCount,
                  False,True);
  TobHisto := Tob.Create('lhistorique',Nil,-1);
  for i := 1 to (GEltDyn.RowCount - 1) do
  begin
    if (PGTypElement = 'ELD') then
    begin
      Salarie := GEltDyn.CellValues[ColSal,i];
      DateAp := StrToDateTime(GEltDyn.CellValues[ColDateV,i]);
      DateFinVal := IDate1900;

      For c := 0 to PGTobLesElements.Detail.Count - 1 do
      begin
        PgTypeDonne := PGTobLesElements.Detail[c].GetValue('PGTYPEDONNE');
        CodTabl := PGTobLesElements.Detail[c].GetValue('CODTABL');
        PgInfosModif := PGTobLesElements.Detail[c].GetValue('PGINFOSMODIF');
        NumChamp := PGTobLesElements.Detail[c].GetValue('NUMCHAMP');
        BStop := False;
        If PgTypeDonne = 'D' then
        begin
          ValApres := GEltDyn.CellValues[1 + (NumChamp*2),i];
          If not IsValidDate(ValApres) then
          begin
            ValApres := DateToStr(IDate1900);
            BStop := True;
          end;
        end
        else
        begin
          ValApres := GEltDyn.CellValues[1 + (NumChamp*2),i];
          if ValApres = '' then
            BStop := True;
        end;

        if not BStop then
        begin
          // Vérifier qu'il n'y a pas déjà une valeur à cette date
          if not ExisteSQL('SELECT PHD_PGINFOSMODIF FROM PGHISTODETAIL ' +
            ' WHERE PHD_SALARIE="' + Salarie + '" AND PHD_DATEAPPLIC="'+ USDATETIME(DateAp) + '" AND PHD_PGINFOSMODIF="' + PgInfosModif + '"') then
          begin
            Q := OpenSQL('SELECT PHD_NEWVALEUR,PHD_DATEAPPLIC FROM PGHISTODETAIL ' +
              ' WHERE PHD_SALARIE="' + Salarie + '" AND PHD_DATEAPPLIC<"'+ USDATETIME(DateAp) + '" AND PHD_PGINFOSMODIF="' + PgInfosModif + '"' +
              ' ORDER BY PHD_DATEAPPLIC DESC',True);
            if (Q.Eof) or (Q.FindField('PHD_NEWVALEUR').AsString <> ValApres) then
            begin
              TH := Tob.Create('PGHISTODETAIL',TobHisto,-1);
              TH.PutValue('PHD_SALARIE',Salarie);
              TH.PutValue('PHD_ETABLISSEMENT','');
              TH.PutValue('PHD_ORDRE',0);
              TH.PutValue('PHD_GUIDHISTO',AglGetGuid());
              TH.PutValue('PHD_PGINFOSMODIF',PgInfosModif);
              If DateFinVal <> IDate1900 then TH.PutValue('PHD_PGTYPEHISTO','004')
              else TH.PutValue('PHD_PGTYPEHISTO','003');
              TH.PutValue('PHD_NEWVALEUR',ValApres);
              TH.PutValue('PHD_TYPEVALEUR',PgTypeDonne);
              TH.PutValue('PHD_TABLETTE','PGCOMBOZONELIBRE');
              TH.PutValue('PHD_CODTABL',CodTabl);
              TH.PutValue('PHD_PGTYPEINFOLS','ZLS');
              TH.PutValue('PHD_DATEAPPLIC',DateAp);
              If TestDate >= DateAp then TH.PutValue('PHD_TRAITEMENTOK','X')
              else TH.PutValue('PHD_TRAITEMENTOK','-');
              TH.PutValue('PHD_DATEFINVALID',DateFinVal);
              TH.PutValue('PHD_TYPEBUDG','');
              TH.PutValue('PHD_NUMAUG',-1);
              TH.PutValue('PHD_ANNEE','');
              TH.InsertDB(Nil,False);
            end
            else
              TraceErr.Items.Add('Pour le Salarié '+Salarie+', la valeur pour le champ '+PgInfosModif+' ' + PGTobLesElements.Detail[c].GetValue('LIBELLE') + ' est identique à la valeur en date du ' + DateTimeToStr(Q.FindField('PHD_DATEAPPLIC').AsDateTime));
            Ferme(Q);
          end
          else
          begin
            Q := OpenSQL('SELECT PHD_NEWVALEUR FROM PGHISTODETAIL ' +
            ' WHERE PHD_SALARIE="' + Salarie + '" AND PHD_DATEAPPLIC="'+ USDATETIME(DateAp) + '" AND PHD_PGINFOSMODIF="' + PgInfosModif + '"',True);
            if (ValApres <> Q.FindField('PHD_NEWVALEUR').AsString) then
              TraceErr.Items.Add('Pour le Salarié '+Salarie+', il existe déjà une valeur pour le champ '+PgInfosModif+' ' + PGTobLesElements.Detail[c].GetValue('LIBELLE') + ' à la date du ' + DateTimeToStr(DateAp));
            Ferme(Q);
          end;
        end;
      end;
    end
    else
    begin
      Salarie := GEltDyn.CellValues[ColSal,i];
      DateAp := StrToDateTime(GEltDyn.CellValues[ColDateV,i]);
      DateFinVal := IDate1900;
      For c := 0 to PGTobLesElements.Detail.Count - 1 do
      begin
        CodeElt := PGTobLesElements.Detail[c].GetValue('CODEELT');
        NumChamp := PGTobLesElements.Detail[c].GetValue('NUMCHAMP');
        ValApres := FloatToStr(FormatageDouble(GEltDyn.CellValues[1 + (NumChamp*2),i]));
        If not IsNumeric(ValApres) then
          ValApres := '0';

        if not ExisteSQL('SELECT PED_CODEELT FROM ELTNATIONDOS ' +
          ' WHERE PED_TYPENIVEAU="SAL" AND PED_VALEURNIVEAU="' + Salarie + '"' +
          ' AND PED_CODEELT="' + CodeElt + '" AND PED_DATEVALIDITE="' + UsDateTime(DateAp) + '"') then
        begin
          //Récupérer la dernière valeur connue au niveau salarié ou supérieur
          BTrouve := False;
          Etab := '';
          ConvSal := '';
          DateValidite := iDate1900;
          TypeNiveau := '';
          Q := OpenSQL('SELECT PSA_ETABLISSEMENT,PSA_CONVENTION FROM SALARIES WHERE PSA_SALARIE ="'+Salarie+'"',True);
          if not Q.Eof then
          begin
            Etab := Q.FindField('PSA_ETABLISSEMENT').AsString;
            ConvSal := Q.FindField('PSA_CONVENTION').AsString;
          end;
          Ferme(Q);
          Popul := '';
          ValComp := 0;
    {$IFNDEF CPS1}
          Qt := OpenSQL('SELECT PNA_POPULATION FROM SALARIEPOPUL '
            + ' WHERE PNA_SALARIE = "' + Salarie + '"'
            + ' AND PNA_TYPEPOP = "PAI"', True);
          if not Qt.Eof then
            Popul := Qt.FindField('PNA_POPULATION').AsString;
          Ferme(Qt);
    {$ENDIF}

          LocalTob := Tob.Create('ELTNATIONDOS', nil, -1);
          St := ' SELECT PED_CODEELT,PED_TYPENIVEAU,PED_MONTANTEURO,PED_DATEVALIDITE FROM ELTNATIONDOS ' +
            ' WHERE PED_CODEELT="' + CodeElt + '" AND PED_DATEVALIDITE<="' + USDATETIME(DEBUTDEMOIS(V_PGI.DateEntree)) + '"' +
            ' AND ((PED_TYPENIVEAU="SAL" AND PED_VALEURNIVEAU="' + Salarie + '")' +
            ' OR (PED_TYPENIVEAU="ETB" AND PED_VALEURNIVEAU="' + Etab + '")' +
            ' OR (PED_TYPENIVEAU="POP" AND PED_VALEURNIVEAU="' + Popul + '"))' +
            ' ORDER BY PED_TYPENIVEAU, PED_DATEVALIDITE DESC';
          LocalTob.LoadDetailDBFromSQL('ELTNATIONDOS', St);

          // Recherche si une valeur existe déjà au niveau SAL
          TFind := LocalTob.FindFirst(['PED_CODEELT', 'PED_TYPENIVEAU'], [CodeElt, 'SAL'], false);
          if TFind <> nil then
          begin
            ValComp := TFind.GetValue('PED_MONTANTEURO');
            DateValidite := TFind.GetValue('PED_DATEVALIDITE');
            TypeNiveau := TFind.GetValue('PED_TYPENIVEAU');
            BTrouve := True;
          end;
          FreeAndNil(TFind);

          // Recherche si une valeur existe déjà au niveau ETB
          if not BTrouve then
          begin
            TFind := LocalTob.FindFirst(['PED_CODEELT', 'PED_TYPENIVEAU'], [CodeElt, 'ETB'], false);
            if TFind <> nil then
            begin
              ValComp := TFind.GetValue('PED_MONTANTEURO');
              DateValidite := TFind.GetValue('PED_DATEVALIDITE');
              TypeNiveau := TFind.GetValue('PED_TYPENIVEAU');
              BTrouve := True;
            end;
            FreeAndNil(TFind);
          end;

    {$IFNDEF CPS1}
          // Recherche si une valeur existe déjà au niveau POP
          if not BTrouve then
          begin
            TFind := LocalTob.FindFirst(['PED_CODEELT', 'PED_TYPENIVEAU'], [CodeElt, 'POP'], false);
            if TFind <> nil then
            begin
              ValComp := TFind.GetValue('PED_MONTANTEURO');
              DateValidite := TFind.GetValue('PED_DATEVALIDITE');
              TypeNiveau := TFind.GetValue('PED_TYPENIVEAU');
              BTrouve := True;
            end;
            FreeAndNil(TFind);
          end;
    {$ENDIF}

          FreeAndNil(LocalTob);

          if not BTrouve then
          begin
            LocalTob := Tob.Create('ELTNATIONAUX', nil, -1);
            St := ' SELECT PEL_CODEELT,PEL_PREDEFINI,PEL_CONVENTION,PEL_MONTANTEURO,PEL_DATEVALIDITE FROM ELTNATIONAUX ' +
              ' WHERE ##PEL_PREDEFINI## PEL_CODEELT="' + CodeElt + '" AND PEL_DATEVALIDITE<="' + USDATETIME(DEBUTDEMOIS(V_PGI.DateEntree)) + '"' +
              ' ORDER BY PEL_PREDEFINI, PEL_DATEVALIDITE DESC, PEL_CONVENTION';
            LocalTob.LoadDetailDBFromSQL('ELTNATIONAUX', St);

            // Recherche si une valeur existe déjà au niveau DOS
            TFind := LocalTob.FindFirst(['PEL_CODEELT', 'PEL_PREDEFINI'], [CodeElt, 'DOS'], false);
            if TFind <> nil then
            begin
              ValComp := TFind.GetValue('PEL_MONTANTEURO');
              DateValidite := TFind.GetValue('PEL_DATEVALIDITE');
              TypeNiveau := TFind.GetValue('PEL_PREDEFINI');
              BTrouve := True;
            end;
            FreeAndNil(TFind);

            // Recherche si une valeur existe déjà au niveau STD CNV
            if not BTrouve then
            begin
              if (ConvSal <> '') then
              begin
                TFind := LocalTob.FindFirst(['PEL_CODEELT', 'PEL_PREDEFINI','PEL_CONVENTION'], [CodeElt, 'STD', ConvSal], false);
                if TFind <> nil then
                begin
                  ValComp := TFind.GetValue('PEL_MONTANTEURO');
                  DateValidite := TFind.GetValue('PEL_DATEVALIDITE');
                  TypeNiveau := TFind.GetValue('PEL_PREDEFINI');
                  BTrouve := True;
                end;
                FreeAndNil(TFind);
              end;
              Ferme(Q);
            end;

            // Recherche si une valeur existe déjà au niveau STD 000
            if not BTrouve then
            begin
              TFind := LocalTob.FindFirst(['PEL_CODEELT', 'PEL_PREDEFINI','PEL_CONVENTION'], [CodeElt, 'STD', '000'], false);
              if TFind <> nil then
              begin
                ValComp := TFind.GetValue('PEL_MONTANTEURO');
                DateValidite := TFind.GetValue('PEL_DATEVALIDITE');
                TypeNiveau := TFind.GetValue('PEL_PREDEFINI');
                BTrouve := True;
              end;
              FreeAndNil(TFind);
            end;

            // Recherche si une valeur existe déjà au niveau CEG
            if not BTrouve then
            begin
              TFind := LocalTob.FindFirst(['PEL_CODEELT', 'PEL_PREDEFINI'], [CodeElt, 'CEG'], false);
              if TFind <> nil then
              begin
                ValComp := TFind.GetValue('PEL_MONTANTEURO');
                DateValidite := TFind.GetValue('PEL_DATEVALIDITE');
                TypeNiveau := TFind.GetValue('PEL_PREDEFINI');
                BTrouve := True;
              end;
              FreeAndNil(TFind);
            end;
            FreeAndNil(LocalTob);
          end;

          if (ValComp <> StrToFloat(ValApres)) then
          begin
            TH := Tob.Create('ELTNATIONDOS',TobHisto,-1);
            TH.PutValue('PED_TYPENIVEAU','SAL');
            TH.PutValue('PED_VALEURNIVEAU',Salarie);
            TH.PutValue('PED_CODEELT',CodeElt);
            TH.PutValue('PED_DATEVALIDITE',DateAp);

            Req := 'SELECT PEL_LIBELLE,PEL_THEMEELT,PEL_MONETAIRE,PEL_ABREGE,PEL_BLOCNOTE,PEL_DECALMOIS,'
              + ' PEL_REGIMEALSACE,PEL_PREDEFINI,PEL_NODOSSIER,PEL_CONVENTION,PEL_MONTANT,PEL_MONTANTEURO FROM ELTNATIONAUX'
              + ' WHERE ##PEL_PREDEFINI## AND PEL_DATEVALIDITE <= "' + UsDateTime(DateAp) + '"'
              + ' AND PEL_CODEELT = "' + CodeElt + '"'
              + ' ORDER BY PEL_DATEVALIDITE DESC';
            Q := OpenSQL(Req,True);
            if not Q.eof then
            begin
              TH.PutValue('PED_LIBELLE', Q.FindField('PEL_LIBELLE').AsString);
              TH.PutValue('PED_THEMEELT', Q.FindField('PEL_THEMEELT').AsString);
              TH.PutValue('PED_MONETAIRE', Q.FindField('PEL_MONETAIRE').AsString);
              TH.PutValue('PED_ABREGE', Q.FindField('PEL_ABREGE').AsString);
              TH.PutValue('PED_BLOCNOTE', Q.FindField('PEL_BLOCNOTE').AsString);
              TH.PutValue('PED_DECALMOIS', Q.FindField('PEL_DECALMOIS').AsString);
              TH.PutValue('PED_REGIMEALSACE', Q.FindField('PEL_REGIMEALSACE').AsString);
            end
            else
            begin
              Ferme(Q);
              Req := 'SELECT PEL_LIBELLE,PEL_THEMEELT,PEL_MONETAIRE,PEL_ABREGE,PEL_BLOCNOTE,PEL_DECALMOIS,'
                + ' PEL_REGIMEALSACE,PEL_PREDEFINI,PEL_NODOSSIER,PEL_CONVENTION,PEL_MONTANT,PEL_MONTANTEURO FROM ELTNATIONAUX'
                + ' WHERE ##PEL_PREDEFINI## AND PEL_CODEELT = "' + CodeElt + '"'
                + ' ORDER BY PEL_DATEVALIDITE DESC';
              Q := OpenSQL(Req,True);
              if not Q.eof then
              begin
                TH.PutValue('PED_LIBELLE', Q.FindField('PEL_LIBELLE').AsString);
                TH.PutValue('PED_THEMEELT', Q.FindField('PEL_THEMEELT').AsString);
                TH.PutValue('PED_MONETAIRE', Q.FindField('PEL_MONETAIRE').AsString);
                TH.PutValue('PED_ABREGE', Q.FindField('PEL_ABREGE').AsString);
                TH.PutValue('PED_BLOCNOTE', Q.FindField('PEL_BLOCNOTE').AsString);
                TH.PutValue('PED_DECALMOIS', Q.FindField('PEL_DECALMOIS').AsString);
                TH.PutValue('PED_REGIMEALSACE', Q.FindField('PEL_REGIMEALSACE').AsString);
              end;
            end;
            Ferme(Q);

            TH.PutValue('PED_MONTANT', 0);
            TH.PutValue('PED_MONTANTEURO', StrToFloat(ValApres));

            TH.PutValue('PED_SALARIE',Salarie);
            TH.PutValue('PED_ETABLISSEMENT','');
            TH.PutValue('PED_CODEPOP','');
            TH.InsertDB(Nil,False);
          end
          else
            TraceErr.Items.Add('Pour le Salarié '+Salarie+', la valeur pour le champ '+PgInfosModif+' ' +
                    PGTobLesElements.Detail[c].GetValue('LIBELLE') + ' est identique à la valeur de niveau ' +
                    TypeNiveau + ' en date du ' + DateTimeToStr(DateValidite));
          Ferme(Qt);
        end
        else
        begin
          Q := OpenSQL('SELECT PED_MONTANTEURO FROM ELTNATIONDOS ' +
            ' WHERE PED_TYPENIVEAU="SAL" AND PED_VALEURNIVEAU="' + Salarie + '"' +
            ' AND PED_CODEELT="' + CodeElt + '" AND PED_DATEVALIDITE="' + UsDateTime(DateAp) + '"',True);
          if (StrToFloat(ValApres) <> Q.FindField('PED_MONTANTEURO').AsFloat) then
            TraceErr.Items.Add('Pour le Salarié '+Salarie+', il existe déjà une valeur différente pour l''élément '+
                    CodeElt+' ' + PGTobLesElements.Detail[c].GetValue('LIBELLE') + ' à la date du ' +
                    DateTimeToStr(DateAp));
          Ferme(Q);
        end;
      end;
    end;
    MoveCurProgressForm('Salarié : ' + Salarie);
  end;
  PGTobEltDyn.GetGridDetail (GEltDyn,GEltDyn.RowCount-1,'',StChampGrid);
  FiniMoveProgressForm;
  TobHisto.Free;

  For i := 0 to PGTobLesElements.Detail.Count - 1 do
  begin
    NumChamp := PGTobLesElements.Detail[i].GetValue('NUMCHAMP');
    ColAp := 1 + (NumChamp * 2);
    if (PGTypElement = 'ELD') then
    begin
      PgTypeDonne := PGTobLesElements.Detail[i].GetValue('PGTYPEDONNE');
      if PgTypeDonne = 'D' then
        GEltDyn.ColFormats[ColAp] := ShortdateFormat;
{      else If PgTypeDonne = 'B' then
      begin
        GEltDyn.colTypes[ColAP] := 'B';
        GEltDyn.ColFormats[ColAp] := IntToStr(Ord(csCheckBox));
        GEltDyn.ColAligns[ColAP] := taCenter;
      end;}
    end
    else
      GEltDyn.ColFormats[ColAp] := '# ##0.000';
  end;
end;

procedure TOF_PGELTDYNGROUPEE.SaisieGlobale(Sender : TObject);
begin
  NextPrevControl(TFFiche(Ecran));
  PGTobEltDyn.GetGridDetail (GEltDyn,GEltDyn.RowCount-1,'',StChampGrid);
  AGLLanceFiche('PAY','ELTDYNGROUPEEGEN','','',GetControlText('DATEEFFET'));
  PGTobEltDyn.PutGridDetail(GEltDyn,False,False,StChampGrid,False);
  GEltDyn.RowCount := PGTobEltDyn.Detail.Count + 1;
end;

procedure TOF_PGELTDYNGROUPEE.GrilleCopierColler (Fichier :String);
var
  T : TOB;
  i,j,x : integer;
  St,Salarie,FileN : string;
  TChamps,TC,TG : Tob;
  Valeur : String;
  Q : TQuery;
  FichierOk : Boolean;
  TraceErr: TListBox;
  PgTypeDonne, CodTabl, PgInfosModif, Libelle : String;
begin
  TraceErr:= TListBox (GetControl ('LSTBXERROR'));
  if TraceErr = NIL then
  begin
    PgiBox ('Attention, composant trace non trouvé, abandon du traitement',Ecran.Caption);
    exit;
  end;
  TraceErr.Refresh ;
  TraceErr.Clear;

  FichierOk := True;
  if GEltDyn = NIL then
  begin
    PgiBox ('Attention, grille non identifiée',Ecran.Caption);
    exit;
  end;

  if Fichier = '' then
  begin
    SourisSablier;
    T := TOBLoadFromClipBoard;
    SourisNormale;
    if T.Detail.Count <= 1 then
    begin
      PgiBox ('Attention, vous n''avez rien sélectionné dans votre feuille EXCEL',Ecran.Caption);
      exit;
    end;
  end
  else
  begin
    FileN := Fichier;
    if FileN = '' then
    begin
      PgiBox ('Attention, vous n''avez pas sélectionné de fichier EXCEL',Ecran.Caption);
      exit;
    end;
    if not FileExists(FileN) then
    begin
      PgiBox ('Attention, votre fichier EXCEL n''existe pas',Ecran.Caption);
      exit;
    end;
    T := TOB.Create('Ma tob',Nil, -1);
    SourisSablier;
    ImportTOBFromXLS (T, FileN);
    SourisNormale;
  end;
  TChamps := Tob.Create('LesChampsExcel',Nil,-1);
  for j := 0 to T.Detail[1].ChampsSup.Count-1 do
  begin
    st := TCS(T.Detail[1].Champssup[j]).Nom;
    st := copy(st,2,Length(st)-1);
    If PGTobLesElements.FindFirst(['PGINFOSMODIF'],[St],False) <> Nil then
    begin
      if (PGTypElement = 'ELD') then
      begin
        Q := OpenSQL('SELECT PPP_PGINFOSMODIF, PPP_PGTYPEDONNE, PPP_CODTABL, PPP_LIBELLE, PPP_LIENASSOC FROM PARAMSALARIE ' +
          ' WHERE ##PPP_PREDEFINI## PPP_PGTYPEINFOLS="ZLS" AND PPP_TYPENIVEAU="SAL" AND PPP_PGINFOSMODIF="' + St + '"',True);
        If Not Q.Eof then
        begin
          PgTypeDonne := Q.FindField('PPP_PGTYPEDONNE').AsString;
          CodTabl := Q.FindField('PPP_CODTABL').AsString;
          PgInfosModif := Q.FindField('PPP_PGINFOSMODIF').AsString;
          Libelle := Q.FindField('PPP_LIBELLE').AsString;
          TC := Tob.Create('FilleExcel',TChamps,-1);
          TC.AddChampSupValeur('PGINFOSMODIF',PgInfosModif,False);
          TC.AddChampSupValeur('PGTYPEDONNE',PgTypeDonne,False);
          TC.AddChampSupValeur('CODTABL',CodTabl,False);
          TC.AddChampSupValeur('LIBELLE',Libelle,False);
        end
        else
        begin
          PgTypeDonne := '';
          CodTabl := '';
          PgInfosModif := '';
        end;
      end
      else
      begin

      end;
      Ferme(Q);
    end;
  end;
  InitMoveProgressForm (NIL,'Traitement en cours', 'Veuillez patienter SVP ...',T.Detail.Count,FALSE,TRUE);
  For i := 0 to T.Detail.Count - 1 do
  begin
    Salarie := T.Detail[i].GetValue('MATRICULE');
    if (VH_Paie.PgTypeNumSal='NUM') and
       (length(Salarie)<11) and
       (isnumeric(Salarie)) then
      Salarie:= ColleZeroDevant(StrToInt(Salarie), 10);
    MoveCurProgressForm ('Salarié : '+Salarie+' en cours de traitement');
    TG := PGTobEltDyn.FindFirst(['SALARIE'],[Salarie],False);
    If TG <> Nil then
    begin
      For x := 0 to TChamps.Detail.Count - 1 do
      begin
        if (PGTypElement = 'ELD') then
        begin
          St := '_' + TChamps.Detail[x].GetValue('PGINFOSMODIF');
          PgInfosModif := TChamps.Detail[x].GetValue('PGINFOSMODIF');
          Valeur := T.Detail[i].GetValue(St);
          PgTypeDonne := TChamps.Detail[x].GetValue('PGTYPEDONNE');
          Libelle := TChamps.Detail[x].GetValue('LIBELLE');
          If PgTypeDonne = 'S' then
          begin
            TG.PutValue(PgInfosModif,Valeur)
          end
          Else If PgTypeDonne = 'T' then
          begin
            CodTabl := TChamps.Detail[x].GetValue('CODTABL');
            if ExisteSQL(' SELECT PTD_VALCRIT1,PTD_LIBELLECODE,PTD_RESULTAT FROM TABLEDIMDET' +
              ' WHERE PTD_CODTABL="' + CodTabl + '" AND PTD_VALCRIT1 = "' + Valeur + '"') then
              TG.PutValue(PgInfosModif,Valeur)
            else
            begin
              FichierOk := False;
              TraceErr.Items.Add('Pour le Salarié '+Salarie+', la valeur ' + Valeur + ' du champ '+copy(st,2,Length(st)-1)+' ' + Libelle + ' n''existe pas dans le paramétrage');
            end;
          end
          Else If PgTypeDonne = 'D' then
          begin
            If IsValidDate(Valeur) then
              TG.PutValue(PgInfosModif,StrToDate(Valeur))
            else
            begin
              FichierOk := False;
              TraceErr.Items.Add('Pour le Salarié '+Salarie+', la valeur ' + Valeur + ' du champ '+copy(st,2,Length(st)-1)+' ' + Libelle + ' n''est pas une date correcte');
            end;
          end
          Else If PgTypeDonne = 'B' then
          begin
            If (Valeur = 'X') or (Valeur = '-') then
              TG.PutValue(PgInfosModif,Valeur)
            else
            begin
              FichierOk := False;
              TraceErr.Items.Add('Pour le Salarié '+Salarie+', la valeur ' + Valeur + ' du champ '+copy(st,2,Length(st)-1)+' ' + Libelle + ' n''est pas X ou -');
            end;
          end;
        end
        else
        begin
          If PgTypeDonne = 'F' then
          begin
            If (Valeur <> '') and (IsNumeric(Valeur)) then
              TG.PutValue(PgInfosModif,MyStrToFloat(Valeur))
            else
            begin
              FichierOk := False;
              TraceErr.Items.Add('Pour le Salarié '+Salarie+', la valeur ' + Valeur + ' du champ '+copy(st,2,Length(st)-1)+' ' + Libelle + ' n''est pas numérique');
            end;
          end;
        end;
      end;
    end
    else
    begin
      TraceErr.Items.Add('Le Salarié '+Salarie+' est inconnu');
    end;
  end;
  FiniMoveProgressForm();
  If FichierOk then PGTobEltDyn.PutGridDetail(GEltDyn,False,False,StChampGrid,False)
  else PGTobEltDyn.GetGridDetail (GEltDyn,GEltDyn.RowCount-1,'',StChampGrid);
  GEltDyn.RowCount := PGTobEltDyn.Detail.Count + 1;
  If TChamps <> Nil then TChamps.Free;
  If T <> Nil then T.Free;
end;

procedure TOF_PGELTDYNGROUPEE.BFichierClick(Sender : TObject);
var
  Fichier : String;
begin
  Fichier := AGLLanceFiche('PAY','HISTOGROUPEEXLS','','','');
  If Fichier <> '' then GrilleCopierColler(Fichier);
end;

procedure TOF_PGELTDYNGROUPEE.BLegendeClick(Sender : TObject);
begin
  AGLLanceFiche('PAY','HISTOGROUPEELEG','','','ELTDYN_GRP');
end;

procedure TOF_PGELTDYNGROUPEE.KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key=86) and (ssCtrl in Shift) then
    GrilleCopierColler ('');
end;

procedure TOF_PGELTDYNGROUPEE.GrilleColEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var
  T : Tob;
  NumChamp : Double;
  LeType : String;
begin
  if (PGTypElement = 'ELD') then
  begin
    NumChamp := (Ou - 1)/2;
    T := PGTobLesElements.FindFirst(['NUMCHAMP'],[NumChamp],False);
    if T <> Nil then
       LeType := T.GetValue('PGTYPEDONNE');
    if (LeType = 'D') or (LeType = 'T') or (T = nil) Then      //PT1
      GEltDyn.ElipsisButton:=True
    else
      GEltDyn.ElipsisButton:=False;
//PT1    if (LeType = 'T') Then
//PT1      GEltDyn.ColFormats[Ou] := 'CB=PGCOMBOZONELIBRE|PTD_CODTABL="' + T.GetValue('CODTABL') +'" AND PTD_DTVALID=(SELECT MAX(PTD_DTVALID) FROM TABLEDIMDET WHERE PTD_CODTABL="' + T.GetValue('CODTABL') + '" AND PTD_DTVALID<="' + USDATETIME(StrToDateTime(GEltDyn.CellValues[2,GEltDyn.Row])) + '")';
  end;
end;

Function TOF_PGELTDYNGROUPEE.FormatageDouble(Chaine : String) : Double;
var
  Longueur,Indice : Integer;
  St : String;
begin
  longueur:=Length (Chaine); 
  If Longueur < 1 then
  begin
    Result := 0;
    exit;
  end;
  indice:=1;
  repeat
    if (Chaine [Indice]<>' ') then
    St := St + Chaine [Indice];
    inc (Indice);
  until (Indice=Longueur+1);
  If Not IsNumeric (St) then
  begin
    Result := 0;
    PGIBox(St + ' n''est pas une valeur numérique',Ecran.Caption);
    Exit;
  end;
  result := StrToFloat(St);
end;

Procedure TOF_PGELTDYNGROUPEE.changementDate(Sender : TObject);
var i : Integer;
begin
  For i := 1 to GEltDyn.RowCount - 1 do
  begin
    GEltDyn.Cellvalues[2,i] := GetControlText('DATEEFFET');
  end;
end;

procedure TOF_PGELTDYNGROUPEE.GrilleKeyPress(Sender: TObject; var Key: Char);
var
  Caract : Char;
  Chaine : String;
begin
  Caract := #0;
  if (PGTypElement <> 'ELD') and (GEltDyn.Col > 2) then
  begin
    if Key in ['.',','] then
    begin
      Chaine := GEltDyn.Cells[GEltDyn.col,GEltDyn.Row];
      if pos(DecimalSeparator,Chaine) = 0 then
        Caract := DecimalSeparator;
    end
    else if (Key in ['0'..'9',#8]) then
      Caract := Key;
    Key := Caract;
  end;
end;

{procedure TOF_PGELTDYNGROUPEE.ImprimeListe(Sender: TObject);
var
  TobEdit, TOBE : Tob;
  ListeErr : TListBox;
  i : integer;
  Titre : String;
begin
  ListeErr := TListbox(Getcontrol('LSTBXERROR'));
  if (ListeErr.Count <> 0) then
  begin
    TobEdit := TOB.Create('',nil,-1);
    For i := 0 to ListeErr.Count -1 do
    begin
      TOBE := TOB.Create('',TobEdit,-1);
      TOBE.AddChampSupValeur('LIGNE',ListeErr.Items[i],false);
    end;
    if (PGTypElement = 'ELD') then
      Titre := 'Liste d''erreurs de la saisie groupée des éléments dynamiques'
    else
      Titre := 'Liste d''erreurs de la saisie groupée des éléments dossier';
    LanceEtatTob('E','PAY','GLS',TobEdit,True,False,False,TPageControl(GetControl('Pages')),'',Titre,False);
    FreeAndNil(TobEdit);
  end
  else
    PGIBox('Il n''y a aucune erreur à imprimer');
end; }

Initialization
  registerclasses ( [ TOF_PGELTDYNGROUPEE ] ) ;
end.

