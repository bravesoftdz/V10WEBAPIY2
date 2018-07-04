{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 25/05/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGHISTOGROUPEE ()
Mots clefs ... : TOF;PGHISTOGROUPEE
*****************************************************************
PT1 15/11/2005 JL V_650 Corrections formats double en validation
PT2 21/04/2006 JL V_650 FQ 13090 Affichage plein écran
PT3 11/05/2007 JL V_720 Modification des fonctions de validation pour tenir compte de la nouvelle gestion de l'historique
PT4 25/05/2007 GGU V_720 Mise à jour de la population du salarié lors de la modification de la fiche salarié
PT5 28/05/2007 GGU V_720 FQ 13990 lorsqu'on importe un fichier excel il s'arrête à 2 chiffres après la virgule au lieu de 3
PT6 30/05/2007 JL V_720 Fq 14267 Suop affichage plein écran par défaut
PT7 10/07/2007 MF V_720 FQ 13777  Ajout "unité pris dans l'effectif" dans les
                                  zones de la liste des informations à modifier
PT8 02/08/2007 JL V_80 FQ 14622 Gestion changement de date global
PT9 23/10/2007 GGU V_80 FQ 14880 l'unicité de l'enregistrement n'est pas contrôlée en validation de la saisie
}
Unit UTofPGHistoGroupee ;

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
     UTofPGMulHistoGroupee,
     UTOM,
     ParamDat,
     PGOutils2,
     PGOutilsHistorique,
     UTOF,
     PGPOPULOUTILS ;

Const
     ColSal = 0;
     ColNom = 1;
     ColDateV = 2;
     ColDateF = 3;
     ColEtat = 4;

Type
  TOF_PGHISTOGROUPEE = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose  ; override ;
    procedure OnLoad  ; override ;
    private
    GHistorique : THGrid;
    NbChamps : Integer;
    StChampGrid : String;
    TitreColonne: TStringList;
    UpdateIdemPop : TUpdateIdemPop; //PT4
    procedure MiseEnFormeGrille;
    procedure RemplirGrille;
    procedure GrillePostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
//    procedure GrilleGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
    procedure ZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
    function ZoneAccessible(var ACol, ARow: Longint): Boolean;
    procedure GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GrilleCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
    procedure InsererLigne (Sender : TObject);
    procedure SuppLigne (Sender : TObject);
    procedure ValiderSaisie (Sender : TObject);
    procedure SaisieGlobale(Sender : TObject);
    procedure SaisieMontant(Sender : TObject);
    procedure IntegrationFicheSal(Champ,TypeChamp,Valeur,Salarie : String);
    procedure GrilleCopierColler(Fichier :String);
    procedure BFichierClick(Sender : TObject);
    procedure BLegendeClick(Sender : TObject);
    procedure KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GrilleElipsisClick(Sender: TObject);
    procedure GrilleColEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    Function FormatageDouble(Chaine : String) : Double;
    Function FormatageInt(Chaine : String) : Integer;
    Procedure changementDate(Sender : TObject);
  end ;

Implementation


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

procedure TOF_PGHISTOGROUPEE.OnClose  ;
begin
     Inherited ;
     if TitreColonne <> nil then TitreColonne.Free;
  FreeAndNil(UpdateIdemPop); //PT4
end;

procedure TOF_PGHISTOGROUPEE.OnLoad  ;
begin
     Inherited ;
end;


procedure TOF_PGHISTOGROUPEE.OnArgument (S : String ) ;
var LeChamp : String;
    BLegende,BFichier,BAjout,BMoins,BVal,BSaisie,Bt : TToolBarButton97;
    NumChamp,i : Integer;
    T : Tob;
begin
  Inherited ;
//  TFVierge(Ecran).WindowState := wsMaximized; //PT6
  UpdateIdemPop := TUpdateIdemPop.Create;  //PT4
//  TFVierge(Ecran).Align := AlClient;
  NumChamp := 1;
  LeChamp := ReadTokenpipe(S,';');
  StChampGrid := 'SALARIE;LIBELLE;DATEAPPLIC;DATEFINVAL;ETAT';
{  PGTobLesChamps := Tob.Create('LesChamps',Nil,-1);
  PGTobLesChamps.LoadDetailDB('LesChamps','','',Q,False);
  While LeChamp <> '' do
  begin
       TLC := Tob.Create('UnChamp',PGTobLesChamps,-1);
       Q := OpenSQL('SELECT PAI_PREFIX,PAI_SUFFIX,PAI_LIBELLE,PAI_LIENASSOC,PAI_LETYPE FROM PAIEPARIM WHERE PAI_IDENT="'+LeChamp+'"',True);
       If Not Q.Eof then
       begin
            LeType := Q.FindField('PAI_LETYPE').AsString;
            Tablette := Q.FindField('PAI_LIENASSOC').AsString;
            Prefixe := Q.FindField('PAI_PREFIX').AsString;
            Suffixe := Q.FindField('PAI_SUFFIX').AsString;
            Libelle := Q.FindField('PAI_LIBELLE').AsString;
       end;
       Ferme(Q);
       TLC.AddChampSupValeur('CHAMP',Prefixe+'_'+Suffixe);
       TLC.AddChampSupValeur('TYPE',LeType);
       TLC.AddChampSupValeur('TABLETTE',Tablette);
       TLC.AddChampSupValeur('NUMERO',NumChamp);
       TLC.AddChampSupValeur('LIBCHAMP',Libelle);
       TLC.AddChampSupValeur('IDENT',LeChamp);
       StChampGrid := StChampGrid+';'+Prefixe+'_'+Suffixe;
       StChampGrid := StChampGrid+';'+Prefixe+'_'+Suffixe+'M';
       StChampGrid := StChampGrid+';ORDRE'+IntToStr(NumChamp);
       NumChamp := NumChamp + 1;
       LeChamp := ReadTokenpipe(S,';');
  end;                   }
  NbChamps := PGTobLesChamps.detail.Count;
  For i := 0 to NbChamps - 1 do
  begin
       StChampGrid := StChampGrid+';'+PGTobLesChamps.detail[i].GetValue('CHAMP');
       StChampGrid := StChampGrid+';'+PGTobLesChamps.detail[i].GetValue('CHAMP')+'M';
       StChampGrid := StChampGrid+';ORDRE'+IntToStr(NumChamp);
       NumChamp := NumChamp + 1;
  end;
  GHistorique := THGrid(GetControl('GHISTO'));
  TitreColonne := TStringList.Create;

  GHistorique.OnElipsisClick := GrilleElipsisClick;
  GHistorique.OnColEnter := GrilleColEnter;
  GHistorique.PostDrawCell := GrillePostDrawCell;
  GHistorique.OnCellEnter := GrilleCellEnter;
  GHistorique.OnCellExit := GrilleCellExit;
//  GHistorique.GetCellCanvas := GrilleGetCellCanvas;


  BAjout := TToolBarButton97(GetControl('BAJOUT'));
  BMoins := TToolBarButton97(GetControl('BMOINS'));
  If BAjout <> nil then BAjout.OnClick := InsererLigne;
  If BMoins <> nil then BMoins.OnClick := SuppLigne;
  SetControlVisible('BAJOUT',False);
  SetControlVisible('BMOINS',False);
  BVal := TToolBarButton97(GetControl('BVALIDSAISIE'));
  If BVal <> Nil then BVal.OnClick := ValiderSaisie;
  BSaisie := TToolBarButton97(GetControl('BSAISIE'));
  If BSaisie <> Nil then BSaisie.OnClick := SaisieGlobale;
  BSaisie := TToolBarButton97(GetControl('BNUMERIQUE'));
  If BSaisie <> Nil then BSaisie.OnClick := SaisieMontant;
  BFichier := TToolBarButton97(GetControl('BFICHIER'));
  If BFichier <> Nil then BFichier.OnClick := BFichierClick;
  GHistorique.OnKeyDown := KeyDown;
  BLegende := TToolBarButton97(GetControl('BLEGENDE'));
  If BLegende <> Nil then BLegende.OnClick := BLegendeClick;
  T := PGTobLesChamps.FindFirst(['LETYPE'],['F'],False);
  If T = Nil then
  begin
       T := PGTobLesChamps.FindFirst(['LETYPE'],['I'],False);
       If T = Nil then SetControlenabled('BNUMERIQUE',False);
  end;
  MiseEnFormeGrille;
  RemplirGrille;
  SetControlText('DATEEFFET',DateToStr(DebutDeMois(V_PGI.DateEntree)));
  Bt :=  TToolBarButton97(GetControl('BDATE'));
  If Bt <> Nil then Bt.OnClick := ChangementDate;
end ;

procedure TOF_PGHISTOGROUPEE.MiseEnFormeGrille;
var i,NumChamp : Integer;
    LeChamp,TypeChamp,Tablette,Libelle : String;
    Titres : HTStringList;
    ColAv,ColAp,ColOrdre,NbInfos : Integer;
begin
     TitreColonne.Clear;
     GHistorique.ColCount := 5 + (PGTobLesChamps.Detail.Count)*3;
     Titres := HTStringList.Create;
     GHistorique.ColEditables[ColSal] := False;
     GHistorique.ColFormats[ColSal] := '';
     Titres.Insert(ColSal, 'Salarie');
     TitreColonne.Add('SALARIE');
     GHistorique.ColEditables[ColNom] := False;
     GHistorique.ColFormats[ColNom] := '';
     Titres.Insert(ColNom, 'Nom');
     TitreColonne.Add('LIBELLE');
     GHistorique.ColWidths[ColNom] := 120;
     GHistorique.ColFormats[ColDateV] := ShortDateFormat;
     Titres.Insert(ColDateV, 'Date');
     TitreColonne.Add('DATEVAL');
     GHistorique.ColWidths[ColDateV] := 100;
     GHistorique.ColFormats[ColDateF] := ShortDateFormat;
     Titres.Insert(ColDateF, 'Date de fin');
     TitreColonne.Add('DATEFINVAL');
     GHistorique.ColWidths[ColDateF] := -1;
     GHistorique.ColEditables[ColEtat] := False;
     GHistorique.ColFormats[ColEtat] := '';
     Titres.Insert(ColEtat, 'Etat');
     TitreColonne.Add('ETAT');
     GHistorique.ColWidths[ColEtat] := -1;
     PGTobLesChamps.Detail.Sort('NUMERO');
     NbInfos := PGTobLesChamps.Detail.Count;
     For i := 0 to PGTobLesChamps.Detail.Count - 1 do
     begin
          NumChamp := PGTobLesChamps.Detail[i].GetValue('NUMERO');
          ColAv := 2 + (NumChamp * 3);
          GHistorique.ColEditables[ColAv] := False;
          ColAp := 3 + (NumChamp * 3);
          ColOrdre  := 4 + (NumChamp * 3);
          TypeChamp := PGTobLesChamps.Detail[i].GetValue('LETYPE');
          Tablette := PGTobLesChamps.Detail[i].GetValue('TABLETTE');
          Libelle := PGTobLesChamps.Detail[i].GetValue('LIBCHAMP');
          LeChamp := PGTobLesChamps.Detail[i].GetValue('CHAMP');
          GHistorique.ColWidths[ColOrdre] := -1;
          If NbInfos < 5 then
          begin
               GHistorique.ColWidths[ColAv] := 100;
               GHistorique.ColWidths[ColAp] := 100;
          end
          else
          begin
               GHistorique.ColWidths[ColAv] := 90;
               GHistorique.ColWidths[ColAp] := 90;
          end;
          Titres.Insert(ColAv, Libelle);
          TitreColonne.Add('ANCIEN'+IntToStr(NumChamp));
          Titres.Insert(ColAp, 'Nouveau');
          TitreColonne.Add(LeChamp);
          Titres.Insert(4+ (NumChamp * 3), 'Ordre');
          TitreColonne.Add('ORDRE'+IntToStr(NumChamp));
          if TypeChamp = 'D' then
          begin
               GHistorique.ColFormats[ColAv] := ShortdateFormat;
               GHistorique.ColFormats[ColAp] := ShortdateFormat;
               GHistorique.ColTypes[ColAv] := 'D';
               GHistorique.ColTypes[ColAp] := 'D';
          end
          else If TypeChamp = 'F' then
          begin
               GHistorique.ColFormats[ColAv] := '# ##0.000';  //PT5
               GHistorique.ColFormats[ColAp] := '# ##0.000';  //PT5
          end
          else If TypeChamp = 'I' then
          begin
               GHistorique.ColFormats[ColAv] := '# ##0';
               GHistorique.ColFormats[ColAp] := '# ##0';
          end
          else If TypeChamp = 'B' then
          begin
               GHistorique.colTypes[ColAv] := 'B';
               GHistorique.colTypes[ColAP] := 'B';
               GHistorique.ColFormats[ColAv] := IntToStr (Ord (csCheckBox));
               GHistorique.ColFormats[ColAP] := IntToStr (Ord (csCheckBox));
          end
          else If TypeChamp = 'T' then
          begin
               GHistorique.ColFormats[ColAv] := 'CB='+Tablette;
               GHistorique.ColFormats[ColAp] := 'CB='+Tablette;
          end;
     end;
     GHistorique.Titres := Titres;
     Titres.free;
end;

procedure TOF_PGHISTOGROUPEE.RemplirGrille;
begin
     PGTobHistoSal.PutGridDetail(GHistorique,False,False,StChampGrid,False);
     GHistorique.RowCount := PGTobHistoSal.Detail.Count + 1;
end;

procedure TOF_PGHISTOGROUPEE.GrillePostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var i : integer;
begin
   For i := 1 to NbChamps do
   begin
     If ACol = (2 + (i * 3))  then GridGriseCell(GHistorique, Acol, Arow, Canvas);
   end;
end;

{procedure TOF_PGHISTOGROUPEE.GrilleGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
Var i : Integer;
    valeur : String;
begin
   If ARow = 0 then Exit;
   For i := 1 to NbChamps do
   begin
     If ACol = (3 + (i * 3))  then
     begin
          If Valeur <> GHistorique.CellValues[ACol - 1,ARow] then Canvas.Font.Color := Clred;
     end;
   end;
end;}

procedure TOF_PGHISTOGROUPEE.ZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
var
  Sens, ii: integer;
  OldEna: boolean;
begin
  OldEna := GHistorique.SynEnabled;
  GHistorique.SynEnabled := False;
  Sens := -1;
  if GHistorique.Row > ARow then Sens := 1 else if ((GHistorique.Row = ARow) and (ACol < GHistorique.Col)) then Sens := 1;
  if (sens = -1) and ((ACol = ColSal) or (GHistorique.Col = ColSal)) and (ARow = 1) then
  begin
    GHistorique.SynEnabled := OldEna;
    Cancel := TRUE;
    exit;
  end;
  ACol := GHistorique.Col;
  ARow := GHistorique.Row;
  ii := 0;
  while not ZoneAccessible(ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 1000 then Break;
    if (ACol = ColSal) and (ARow = 1) then
    begin
      ACol := 1;
      break;
    end;
    if Sens = 1 then
    begin
      if ((ACol = GHistorique.ColCount - 1) and (ARow = GHistorique.RowCount - 1)) then
      begin
        ACol := GHistorique.FixedCols;
        ARow := 1;
        Break;
      end;
      if ACol < GHistorique.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        ACol := GHistorique.FixedCols;
      end;
    end else
    begin
      if ((ACol = GHistorique.FixedCols) and (ARow = 1)) then Break;
      if ACol > GHistorique.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := GHistorique.ColCount - 1;
      end;
    end;
  end;
  GHistorique.SynEnabled := OldEna;
end;

function TOF_PGHISTOGROUPEE.ZoneAccessible(var ACol, ARow: Longint): Boolean;
var
  T1: TOB;
  i : integer;
begin
  result := FALSE;
  If (ACol = ColNom) or (ACol = ColDateV) or (ACol = ColDateF) then
  begin
       Result := True;
       Exit;
  end;
  For i := 1 to NbChamps do
  begin
     If ACol = (3 + (i * 3))  then
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
  T1 := TOB(GHistorique.Objects[ColSal, GHistorique.Row]);
  if T1 = nil then
  begin
    result := FALSE;
    exit;
  end;
end;

procedure TOF_PGHISTOGROUPEE.GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  ZoneSuivanteouOk(ACol, ARow, Cancel);
end;

procedure TOF_PGHISTOGROUPEE.GrilleCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var Prefixe,LeChamp,TypeChamp,Valeur : String;
    T : Tob;
begin
     LeChamp := TitreColonne.Strings[ACol];
     Prefixe := ReadTokenPipe(LeChamp,'_');
     If  Prefixe = 'PSA' then
     begin
          TypeChamp := '';
          T := PGTobLesChamps.FindFirst(['CHAMP'],[Prefixe+'_'+LeChamp],false);
          If T <> Nil then   TypeChamp := T.GetValue('LETYPE');
          If TypeChamp <> '' then
          begin
               Valeur := GHIstorique.CellValues[ACol,ARow];
               If TypeChamp = 'D' then
               begin
                    If Valeur = '  /  /    ' then
                    begin
                         Valeur := '';
                         GHIstorique.CellValues[ACol,ARow] := '';
                    end;
                    If Valeur <> '' then
                    begin
                         If Not iSValidDate(Valeur) then
                         begin
                              PGIBox(Valeur + 'n''est pas une date valide');
                              GHIstorique.CellValues[ACol,ARow] := GHIstorique.CellValues[ACol-1,ARow];
                         end;
                    end;
               end;
               If TypeChamp = 'F' then
               begin
                    If Valeur <> '' then
                    begin
                         If Not IsNumeric(Valeur) then
                         begin
                              PGIBox(Valeur + 'n''est pas une valeur numérique correct');
                              GHIstorique.CellValues[ACol,ARow] := GHIstorique.CellValues[ACol-1,ARow];
                         end;
                    end;
               end;
               If TypeChamp = 'I' then
               begin
                    If Valeur <> '' then
                    begin
                         If Not IsNumeric(Valeur) then
                         begin
                              PGIBox(Valeur + 'n''est pas une valeur numérique correct');
                              GHIstorique.CellValues[ACol,ARow] := GHIstorique.CellValues[ACol-1,ARow];
                         end;
                    end;
               end;
          end;
     end;
end;

procedure TOF_PGHISTOGROUPEE.InsererLigne (Sender : TObject);
var Q : TQuery;
    Salarie,LeChamp,LeType : String;
    Newligne,i,NumChamp,ColAv,ColAp,ColOrdre : Integer;
begin
     NextPrevControl(TFFiche(ecran));
     NewLigne := GHistorique.Row+1;
     Salarie := GHistorique.CellValues[ColSal,GHistorique.Row];
     GHistorique.InsertRow(NewLigne);
     GHistorique.CellValues[ColSal,NewLigne] := Salarie;
     GHistorique.CellValues[ColNom,NewLigne] := Rechdom('PGSALARIE',Salarie,False);
     GHistorique.CellValues[ColDateV,NewLigne] := DateToStr(PGDateFinHisto);
     GHistorique.CellValues[ColDateF,NewLigne] := DateToStr(IDate1900);
     GHistorique.CellValues[ColEtat,NewLigne] := '';
     For i := 0 to PGTobLesChamps.Detail.count - 1 do
     begin
          LeChamp := PGTobLesChamps.Detail[i].GetValue('CHAMP');
          LeType := PGTobLesChamps.Detail[i].GetValue('LETYPE');
          NumChamp := PGTobLesChamps.Detail[i].GetValue('NUMERO');
          ColAv := 2 + (NumChamp * 3);
          ColAp := 3 + (NumChamp * 3);
          ColOrdre := 4 + (NumChamp * 3);
          GHistorique.CellValues[ColOrdre,NewLigne] := '0';
          GHistorique.CellValues[4 + (NumChamp * 3),NewLigne] := '0';
          Q := OpenSQL('SELECT '+LeChamp+' FROM SALARIES '+
          'WHERE PSA_SALARIE="'+Salarie+'"',True);
          If Not Q.Eof then
          begin
               If (LeType = 'B') or (LeType='S') or (LeType='T') then
               begin
                    GHistorique.CellValues[ColAv,NewLigne] := Q.FindField(LeChamp).AsString;
                    GHistorique.CellValues[ColAp,NewLigne] := Q.FindField(LeChamp).AsString;
               end
               else If LeType = 'D' then
               begin
                    GHistorique.CellValues[ColAv,NewLigne] := DateToStr(Q.FindField(LeChamp).AsDateTime);
                    GHistorique.CellValues[ColAp,NewLigne] := DateToStr(Q.FindField(LeChamp).AsDateTime);
               end
               else if LeType = 'I' then
               begin
                    GHistorique.CellValues[ColAv,NewLigne] := IntToStr(Q.FindField(LeChamp).AsInteger);
                    GHistorique.CellValues[ColAp,NewLigne] := IntToStr(Q.FindField(LeChamp).AsInteger);
               end
               else If LeType = 'F' then
               begin
                    GHistorique.CellValues[ColAv,NewLigne] := FloatToStr(Q.FindField(LeChamp).AsFloat);
                    GHistorique.CellValues[ColAp,NewLigne] := FloatToStr(Q.FindField(LeChamp).AsFloat);
               end;
          end
          else
          begin
               If (LeType = 'B') or (LeType='S') or (LeType='T') then
               begin
                    GHistorique.CellValues[ColAv,NewLigne] := '';
                    GHistorique.CellValues[ColAp,NewLigne] := '';
               end
               else If LeType = 'D' then
               begin
                    GHistorique.CellValues[ColAv,NewLigne] := DateToStr(IDate1900);
                    GHistorique.CellValues[ColAp,NewLigne] := DateToStr(IDate1900);
               end
               else if (LeType = 'I') or (LeType = 'F') then
               begin
                    GHistorique.CellValues[ColAv,NewLigne] := '0';
                    GHistorique.CellValues[ColAp,NewLigne] := '0';
               end;
          end;
          Ferme(Q);
     end;
end;

procedure TOF_PGHISTOGROUPEE.SuppLigne (Sender : TObject);
begin
     GHistorique.DeleteRow(GHistorique.Row);
end;

procedure TOF_PGHISTOGROUPEE.ValiderSaisie (Sender : TObject);
var i,c : Integer;
    TobHisto,TH,TChamp,TParamHitsto : Tob;
    Salarie : String;
    DateAp,DateFinVal : TDateTime;
    TypeChamp,Tablette,Champ,Indexchamp : String;
    ValAvant,ValApres : String;
    ColAp,NumChamp : Integer;
    TestDate : TDateTime;
    Ident : String;
//    LeChamp,Valeur : String;
    Q : TQuery;
    mess : string; // PT7
    ValBool : boolean; // PT7
    Procedure AddHistorique(DateApplication, DateEntree, DateFinValeur : TDateTime;
      CodeSalarie, NomChamp, TypeDuChamp, NomTablette, ValeurAvant, ValeurApres : String;
      TobHistorique : Tob);
    Begin
      BEGINTRANS;
      try
        { //PT9 On supprime les lignes qui auraient déjà été saisies sur le même jour }
        ExecuteSQL('DELETE FROM PGHISTODETAIL '
                  +'WHERE PHD_DATEAPPLIC="'+UsDateTime(DateApplication)+'" '
                  +'  AND PHD_SALARIE="'+CodeSalarie+'" '
                  +'  AND PHD_PGINFOSMODIF="'+NomChamp+'" '
                  +'  AND PHD_PGTYPEINFOLS="SAL"');
        TH := Tob.Create('PGHISTODETAIL',TobHistorique,-1);
        TH.PutValue('PHD_SALARIE',CodeSalarie);
        TH.PutValue('PHD_ETABLISSEMENT','');
        TH.PutValue('PHD_ORDRE',0);
        TH.PutValue('PHD_GUIDHISTO',AglGetGuid());
        TH.PutValue('PHD_PGINFOSMODIF',NomChamp);
        If DateFinValeur <> IDate1900 then TH.PutValue('PHD_PGTYPEHISTO','004')
        else TH.PutValue('PHD_PGTYPEHISTO','003');
        TH.PutValue('PHD_ANCVALEUR',ValeurAvant);
        TH.PutValue('PHD_NEWVALEUR',ValeurApres);
        TH.PutValue('PHD_TYPEVALEUR',TypeDuChamp);
        TH.PutValue('PHD_TABLETTE',NomTablette);
        TH.PutValue('PHD_PGTYPEINFOLS','SAL');
        TH.PutValue('PHD_DATEAPPLIC',DateApplication);
        If DateEntree >= DateApplication then TH.PutValue('PHD_TRAITEMENTOK','X')
        else TH.PutValue('PHD_TRAITEMENTOK','-');
        TH.PutValue('PHD_DATEFINVALID',DateFinValeur);
        TH.PutValue('PHD_TYPEBUDG','');
        TH.PutValue('PHD_NUMAUG',-1);
        TH.PutValue('PHD_ANNEE','');
        TH.InsertDB(Nil,False);
        COMMITTRANS;
      except
        ROLLBACK;
        PGIError('Une erreur est survenue lors de la mise à jour de l''historique salarié.#10#13La modification du champ '+NomChamp+' n''a pas été enregistrée.');
      end;
    End;
begin
     mess := ''; // PT7
     TestDate := V_PGI.DateEntree;
     NextPrevControl(TFFiche(Ecran));
{     For i := 0 to PGTobLesChamps.Detail.count - 1 do
     begin
          NumChamp := PGTobLesChamps.Detail[i].GetValue('NUMERO');
          GHistorique.ColFormats[3 + (NumChamp * 3)] := '';
     end;}
     PGTobHistoSal.GetGridDetail (GHistorique,GHistorique.RowCount-1,'',StChampGrid);
     InitMoveProgressForm (NIL,'Sauvegarde en cours',
                    'Veuillez patienter SVP ...', PGTobHistoSal.Detail.Count,
                    False,True);
     TobHisto := Tob.Create('lhistorique',Nil,-1);
     Q := OpenSQL('SELECT DISTINCT(PPP_PGINFOSMODIF) CHAMP FROM PARAMSALARIE WHERE PPP_HISTORIQUE="X"',True);
     TParamHitsto := Tob.Create('MaTob',Nil,-1);
     TParamHitsto.LoadDetailDB('Table','','',Q,False);
     Ferme(Q);
     For i := 0 to PGTobHistoSal.detail.Count - 1 do
     begin
          Salarie := PGTobHistoSal.detail[i].GetValue('SALARIE');
          DateAp := PGTobHistoSal.detail[i].GetDateTime('DATEAPPLIC');
          If IsValidDate( PGTobHistoSal.detail[i].GetString('DATEFINVAL')) then DateFinVal := PGTobHistoSal.detail[i].GetDateTime('DATEFINVAL')
          else DateFinVal := IDate1900;
          For c := 0 to PGTobLesChamps.Detail.Count - 1 do
          begin
               TypeChamp := PGTobLesChamps.Detail[c].GetValue('LETYPE');
               Tablette := PGTobLesChamps.Detail[c].GetValue('TABLETTE');
               Champ := PGTobLesChamps.Detail[c].GetValue('CHAMP');
               Indexchamp := PGTobLesChamps.Detail[c].GetValue('IDENT');
               TChamp := PGTobLesChamps.FindFirst(['IDENT'],[IDent],False);
               If TypeChamp = 'D' then
               begin
                    ValAvant := PGTobHistoSal.detail[i].GetString(Champ);
                    ValApres := PGTobHistoSal.detail[i].GetString(Champ+'M');
                    If not IsValidDate(ValAvant) then ValAvant := DateToStr(IDate1900);
                    If not IsValidDate(ValApres) then ValApres := DateToStr(IDate1900);
               end
               else If TypeChamp = 'F' then
               begin
                    ValAvant := FloatToStr(FormatageDouble(PGTobHistoSal.detail[i].GetString(Champ))); //PT1
                    ValApres := FloatToStr(FormatageDouble(PGTobHistoSal.detail[i].GetString(Champ+'M')));
                    If not IsNumeric(ValApres) then ValApres := '0';
               end
               else If TypeChamp = 'I' then
               begin
                    ValAvant := FloatToStr(FormatageInt(PGTobHistoSal.detail[i].GetString(Champ))); //PT1
                    ValApres := FloatToStr(FormatageInt(PGTobHistoSal.detail[i].GetString(Champ+'M')));
                    If not IsNumeric(ValApres) then ValApres := '0';
               end
                 else
               begin
                    ValApres := PGTobHistoSal.detail[i].GetString(Champ+'M');
                    ValAvant := PGTobHistoSal.detail[i].GetString(Champ);
               end;
               If ValAvant <> ValApres then
               begin
//                         If (ValApres > DateDePaie) and (DateDePaie <> IDate1900) then MessDate := 'Pour le salarié ..........';

                 If TParamHitsto.FindFirst(['CHAMP'],[Champ],False) <> Nil then
                    AddHistorique(DateAp, TestDate, DateFinVal, Salarie, Champ, TypeChamp, Tablette, ValAvant, ValApres, TobHisto);
                 If TestDate >= DateAp then IntegrationFicheSal(Champ,TypeChamp,ValApres,Salarie);
// d PT7
                 if (Champ = 'PSA_UNITEPRISEFF') then
                 // qd modif du champ PSA_UNITEPRISEFF, modif du champ
                 // PSA_PRISEFFECTIF si nécessaire
                 begin
                    ValBool := false;
                    if (ValAvant = '0') and (ValApres <> '0') then
                    // PSA_PRISEFFECTIF est forcé à True
                    begin
                       ValApres := 'X';
                       Q :=  OpenSql('SELECT PSA_PRISEFFECTIF FROM SALARIES '+
                                     'WHERE PSA_SALARIE ="'+Salarie+'"',True);
                       if not Q.eof then
                          ValBool := Q.FindField('PSA_PRISEFFECTIF').AsString='X' ;
                          if (ValBool) then
                            ValAvant := 'X'
                          else
                            ValAvant := '-';
                       ferme (Q);
                       mess := 'Les salariés modifiés sont pris dans l''effectif';
                    end;
                    if (ValApres = '0') then
                    // PSA_PRISEFFECTIF est forcé à False
                    begin
                       ValApres := '-';
                       Q :=  OpenSql('SELECT PSA_PRISEFFECTIF FROM SALARIES '+
                                     'WHERE PSA_SALARIE ="'+Salarie+'"',True);
                       if not Q.eof then
                          ValBool := Q.FindField('PSA_PRISEFFECTIF').AsString='X';
                          if (ValBool) then
                            ValAvant := 'X'
                          else
                            ValAvant := '-';
                      ferme (Q);
                       mess := 'Les salariés modifiés ne sont plus pris dans l''effectif';
                    end;
                    if (ValApres = 'X') or (ValApres='-') then
                    // historisation de la nouvelle valeur de PSA_PRISEFFECTIF
                    begin
                       AddHistorique(DateAp, TestDate, DateFinVal, Salarie, 'PSA_PRISEFFECTIF', TypeChamp, Tablette, ValAvant, ValApres, TobHisto);
                       If TestDate >= DateAp then IntegrationFicheSal('PSA_PRISEFFECTIF','B',ValApres,Salarie);
                    end;
                 end;
// f PT7
               end;
          end;
          MoveCurProgressForm ('Salarié : '+ Salarie);
     end;
     TParamHitsto.Free;
     FiniMoveProgressForm;
     TobHisto.Free;
     PGTobHistoSal.PutGridDetail(GHistorique,False,False,StChampGrid,False);
     GHistorique.RowCount := PGTobHistoSal.Detail.Count + 1;
     For i := 0 to PGTobLesChamps.Detail.Count - 1 do
     begin
          Tablette := PGTobLesChamps.Detail[i].GetValue('TABLETTE');
          NumChamp := PGTobLesChamps.Detail[i].GetValue('NUMERO');
          ColAp := 3 + (NumChamp * 3);
          TypeChamp := PGTobLesChamps.Detail[i].GetValue('LETYPE');
          if TypeChamp = 'D' then GHistorique.ColFormats[ColAp] := ShortdateFormat
          else If TypeChamp = 'F' then GHistorique.ColFormats[ColAp] := '# ##0.000'  //PT5
          else If TypeChamp = 'I' then GHistorique.ColFormats[ColAp] := '# ##0'
          else If TypeChamp = 'B' then
          begin
               GHistorique.colTypes[ColAP] := 'B';
               GHistorique.ColFormats[ColAp] := IntToStr (Ord (csCheckBox));
          end
          else If TypeChamp = 'T' then GHistorique.ColFormats[ColAp] := 'CB='+Tablette;
     end;
     if (mess <> '') then PgiBox(mess,Ecran.Caption); //PT7
     mess := ''; // PT7

end;

procedure TOF_PGHISTOGROUPEE.SaisieGlobale(Sender : TObject);
begin
     NextPrevControl(TFFiche(Ecran));
     PGTobHistoSal.GetGridDetail (GHistorique,GHistorique.RowCount-1,'',StChampGrid);
     AGLLanceFiche('PAY','HISTOGROUPEEGEN','','','');
     PGTobHistoSal.PutGridDetail(GHistorique,False,False,StChampGrid,False);
     GHistorique.RowCount := PGTobHistoSal.Detail.Count + 1;
end;


procedure TOF_PGHISTOGROUPEE.SaisieMontant(Sender : TObject);
begin
     NextPrevControl(TFFiche(Ecran));
     PGTobHistoSal.GetGridDetail (GHistorique,GHistorique.RowCount-1,'',StChampGrid);
     AGLLanceFiche('PAY','HISTOGROUPEENUM','','','');
     PGTobHistoSal.PutGridDetail(GHistorique,False,False,StChampGrid,False);
     GHistorique.RowCount := PGTobHistoSal.Detail.Count + 1;
end;

procedure TOF_PGHISTOGROUPEE.IntegrationFicheSal(Champ,TypeChamp,Valeur,Salarie : String);
var ValeurMaj : String;
    TobSal : Tob;
    Q : TQuery;
    TmpStringList : TStringList;  //PT4
begin
   ValeurMaj := Valeur;
   If TypeChamp = 'D' then
   begin
    If not IsValidDate(ValeurMaj) then ValeurMaj := DateToStr(IDate1900);
   end
   else If TypeChamp = 'F' then
   begin
    If not IsNumeric(ValeurMaj) then ValeurMaj := '0';
   end
   else If TypeChamp = 'I' then
   begin
    If not IsNumeric(ValeurMaj) then ValeurMaj := '0';
   end;
   Q := OpenSQL('SELECT * FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"',True);
   TobSal := Tob.Create('SALARIES',Nil,-1);
   TobSal.LoadDetailDB('SALARIES','','',Q,False);
   Ferme(Q);
   If TypeChamp = 'D' then TobSal.Detail[0].PutValue(Champ,StrToDate(Valeur))
   else If TypeChamp = 'F' then TobSal.Detail[0].PutValue(Champ,StrToFloat(Valeur))
   else If TypeChamp = 'I' then TobSal.Detail[0].PutValue(Champ,StrToInt(Valeur))
   else TobSal.Detail[0].PutValue(Champ,Valeur);
   TobSal.Detail[0].UpdateDB;
   TobSal.Free;

  //Début PT4 : Mise à jour de la population du salarié lors de la mise à jour de l'historique
  TmpStringList := UpdateIdemPop.MajSALARIEPOPULSalarie(Salarie,V_PGI.DateEntree,[Champ]);
  FreeAndNil(TmpStringList);
  //Fin PT4

end;

procedure TOF_PGHISTOGROUPEE.GrilleCopierColler (Fichier :String);
var T               : TOB;
    i,j,x   : integer;
    St,Salarie,FileN           : string;
    TChamps,TC,TG : Tob;
    Tablette,LeType,Valeur,Colonne : String;
    Q : TQuery;
    FichierOk : Boolean;
    TraceErr: TListBox;
    Suffix : String;
begin
TraceErr:= TListBox (GetControl ('LSTBXERROR'));
if TraceErr = NIL then
     begin
     PgiBox ('Attention, composant trace non trouvé, abandon du traitement',Ecran.Caption);
     exit;
     end;
  TraceErr.Refresh ;
FichierOk := True;
if GHistorique = NIL then
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
  if NOT FileExists(FileN) then
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
     st :=  TCS(T.Detail[1].Champssup[j]).Nom;
     If PGTobLesChamps.FindFirst(['COLONNE'],[St],False) <> Nil then
     begin
          Q := OpenSQL('SELECT PAI_LETYPE,PAI_LIENASSOC,PAI_SUFFIX,PAI_COLONNE FROM PAIEPARIM WHERE PAI_COLONNE="'+St+'"',True);
          If Not Q.Eof then
          begin
               LeType := Q.FindField('PAI_LETYPE').AsString;
               Tablette := Q.FindField('PAI_LIENASSOC').AsString;
               Suffix := Q.FindField('PAI_SUFFIX').AsString;
               Colonne := Q.FindField('PAI_COLONNE').AsString;
               TC := Tob.Create('FilleExcel',TChamps,-1);
               TC.AddChampSupValeur('CHAMP',Suffix,False);
               TC.AddChampSupValeur('LETYPE',LeType,False);
               TC.AddChampSupValeur('TABLETTE',Tablette,False);
               TC.AddChampSupValeur('COLONNE',Colonne,False);
          end
          else
          begin
               LeType := '';
               Tablette := '';
               Suffix := '';
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
     TG := PGTobHistoSal.FindFirst(['SALARIE'],[Salarie],False);
     If TG <> Nil then
     begin
          For x := 0 to TChamps.Detail.Count - 1 do
          begin
               St := TChamps.Detail[x].GetValue('COLONNE');
               Suffix := TChamps.Detail[x].GetValue('CHAMP');
               Valeur := T.Detail[i].GetValue(St);
               LeType := TChamps.Detail[x].GetValue('LETYPE');
               If LeType = 'S' then
               begin
                    TG.PutValue('PSA_'+Suffix+'M',Valeur)
               end
               Else If LeType = 'T' then
               begin
                    Tablette := TChamps.Detail[x].GetValue('TABLETTE');
                    If RechDom(Tablette,Valeur,False) <> '' then TG.PutValue('PSA_'+Suffix+'M',Valeur)
                    else
                    begin
                         FichierOk := False;
                         TraceErr.Items.Add('Pour le Salarié '+Salarie+' le champ '+St+' n''existe pas dans la liste');
                    end;
               end
               Else If LeType = 'D' then
               begin
                    If IsValidDate(Valeur) then TG.PutValue('PSA_'+Suffix+'M',StrToDate(Valeur))
                    else
                    begin
                         FichierOk := False;
                         TraceErr.Items.Add('Pour le Salarié '+Salarie+' le champ '+St+' n''est pas une date correcte');
                    end;
               end
               Else If LeType = 'I' then
               begin
                    If (Valeur <> '') and (IsNumeric(Valeur)) then TG.PutValue('PSA_'+Suffix+'M',StrToInt(Valeur))
                    else
                    begin
                         FichierOk := False;
                         TraceErr.Items.Add('Pour le Salarié '+Salarie+' le champ '+St+' n''est pas numérique');
                    end;
               end
               Else If LeType = 'F' then
               begin
                    If (Valeur <> '') and (IsNumeric(Valeur)) then TG.PutValue('PSA_'+Suffix+'M',MyStrToFloat(Valeur))  //Fonction de transformation . - > , ou inverssement....
                    else
                    begin
                         FichierOk := False;
                         TraceErr.Items.Add('Pour le Salarié '+Salarie+' le champ '+St+' n''est pas numérique');
                    end;
               end
               Else If LeType = 'B' then
               begin
                    If (Valeur = 'X') or (Valeur = '-') then TG.PutValue('PSA_'+Suffix+'M',Valeur)
                    else
                    begin
                         FichierOk := False;
                         TraceErr.Items.Add('Pour le Salarié '+Salarie+' le champ '+St+' n''est pas X ou -');
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
If FichierOk then PGTobHistoSal.PutGridDetail(GHistorique,False,False,StChampGrid,False)
else PGTobHistoSal.GetGridDetail (GHistorique,GHistorique.RowCount-1,'',StChampGrid);
GHistorique.RowCount := PGTobHistoSal.Detail.Count + 1;
If TChamps <> Nil then TChamps.Free;
If T <> Nil then T.Free;
end;

procedure TOF_PGHISTOGROUPEE.BFichierClick(Sender : TObject);
var Fichier : String;
begin
     Fichier := AGLLanceFiche('PAY','HISTOGROUPEEXLS','','','');
     If Fichier <> '' then GrilleCopierColler(Fichier);
end;

procedure TOF_PGHISTOGROUPEE.BLegendeClick(Sender : TObject);
begin
     AGLLanceFiche('PAY','HISTOGROUPEELEG','','','');
end;

procedure TOF_PGHISTOGROUPEE.KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key=86) and (ssCtrl in Shift) then
    GrilleCopierColler ('');
end;

procedure TOF_PGHISTOGROUPEE.GrilleElipsisClick(Sender: TObject);
var key : char;
begin
   key := '*';
   ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGHISTOGROUPEE.GrilleColEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var T : Tob;
    NumChamp : Double;
    LeType : String;
begin
     NumChamp := (Ou - 3)/3;
     T := PGTobLesChamps.FindFirst(['NUMERO'],[NumChamp],False);
     If T <> Nil then LeType := T.GetValue('LETYPE')
     else LeType := '';
     If LeType = 'D' Then GHistorique.ElipsisButton:=True
     Else GHistorique.ElipsisButton:=False;
end;

Function TOF_PGHISTOGROUPEE.FormatageDouble(Chaine : String) : Double;
var Longueur,Indice : Integer;
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

Function TOF_PGHISTOGROUPEE.FormatageInt(Chaine : String) : Integer;
var Longueur,Indice : Integer;
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
        result := StrToInt(St);
end;

Procedure TOF_PGHISTOGROUPEE.changementDate(Sender : TObject); //PT8
var i : Integer;
begin
  For i := 1 to GHistorique.RowCount - 1 do
  begin
    GHistorique.Cellvalues[2,i] := GetControlText('DATEEFFET');
  end;
end;


Initialization
  registerclasses ( [ TOF_PGHISTOGROUPEE ] ) ;
end.

