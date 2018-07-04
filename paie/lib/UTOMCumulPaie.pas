{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Unit de gestion des cumuls de la paie
Mots clefs ... : PAIE
*****************************************************************
PT1 14/11/2001 SB V562 Test suppression cumul existant ds gestion associée
PT2 10/04/2002 Ph V571 Accepttaion des cumuls pairs > 50 à STD
PT3 23/09/2002 SB V585 FQ n°10234 Le code cumul est accessible aprés validation
PT4 18/02/2003 SB V595 FQ 10531 Anomalie sur controle existence rubrique dans table commune
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
PT5 04/10/2004 PH V_50 Suppression rechargement tablette
PT6 03/02/2005 PH V_602 FQ 11337 1ere posistion du cumul en alpha pour augmenter le nbre de cumuls dispo
PT7 26/07/2005 PH V_603 FQ 12440 Test en CWAS sur les controles de cohérence du code du cumul
PT8 14/06/2006 PH V_650 FQ 13287 test longueur du code cumul
PT9 25/01/2007 GGS V_80 FQ 12694  Journal événement
PT9-2 26/04/2007 GGS V_80 Revu gestion Trace
PT10 : 07/05/2007 PH V_72 Concept Paie
PT11 31/05/2007 FC V_72 FQ 14288 Remplacer la codification alphanum minuscule en majuscule
PT12 06/11/2007 FC V_80 Correction bug journal des évènement quand création (lib Modification au lieu de Création)
}
unit UTOMCumulPaie;

interface

uses Classes, HMsgBox,
{$IFNDEF EAGLCLIENT}
  db, FichList,
{$ELSE}
  eFichList, MaineAGL, UtileAGL, Utob,
{$ENDIF}
  UTOM,  HEnt1, SysUtils,
  PgOutils, PgOutils2,
  P5Def,
  PAIETOM,strutils;               //PT9     //PT11


type
  TOM_CUMULPAIE = class(PGTOM)   //PT9 class TOM devient PGTOM
    procedure OnArgument(stArgument: string); override;
    procedure OnChangeField(F: TField); override;
    procedure OnNewRecord; override;
    procedure OnLoadRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnAfterUpdateRecord; override;                   //PT9
{PT9-2    procedure OnArgument(stArgument: string); override;       //PT9}
  private
    LectureSeule, CEG, STD, DOS, Error: Boolean;
    Trace: TStringList;                                       //PT9
    DerniereCreate: string;
    Maj : Boolean;
    LeStatut:TDataSetState; //PT12
  end;
implementation

{ TOM_CUMULPAIE }
//PT9
{PT9-2 procedure TOM_CUMULPAIE.OnArgument;
begin
  Trace := TStringList.Create ;
end;
//FIN PT9
}
procedure TOM_CUMULPAIE.OnChangeField(F: TField);
var
  OKRub: Boolean;
  Cumul, TempCum,TempCum2, vide, mes, Pred: string;
  iCode: integer;
begin
  inherited;
  vide := '';
  if Not Maj then
  begin
    if (F.FieldName = 'PCL_CUMULPAIE') then
    begin
  // DEB PT6
      Cumul := GetField('PCL_CUMULPAIE');
      Pred := GetField('PCL_PREDEFINI');
      if (Cumul = '') OR (Cumul = ' ') OR (Cumul = '  ') then exit; // PT7
      if length(Cumul) < 2 then
      begin
        SetField('PCL_CUMULPAIE', '   ');
        PGIbox('Le code cumul doit comporter 2 caractères', Ecran.Caption);
        SetFocusControl('PCL_CUMULPAIE');
        exit;
      end;

      //DEB PT11 
      LastErrorMsg := '';
      Cumul := GetField('PCL_CUMULPAIE');
      if (not isnumeric(Cumul)) then
      begin
        TempCum := Copy(Cumul, 1, 1);
        if not isnumeric(TempCum) then
        begin
          if (TempCum <> UpperCase(TempCum)) then
          begin
            TempCum := UpperCase(TempCum);
            TempCum2 := Copy(Cumul, 2, 1);
            Cumul := TempCum + TempCum2;
            LastErrorMsg := 'La première lettre du cumul doit être en majuscule'#13#10;
          end;
          if not AnsiContainsText('ABCDEFGHIJKLMNOPQRSTUVWXYZ', TempCum) then
            LastErrorMsg := LastErrorMsg + 'Le premier caractère du cumul doit être une lettre majuscule ou un chiffre de 0 à 9'#13#10;
        end;
        TempCum := Copy(Cumul, 2, 1);
        if not isnumeric(TempCum) then
        begin
          if (TempCum <> UpperCase(TempCum)) then
          begin
            TempCum := UpperCase(TempCum);
            TempCum2 := Copy(Cumul, 1, 1);
            Cumul := TempCum2 + TempCum;
            LastErrorMsg := LastErrorMsg + 'La deuxième lettre du cumul doit être en majuscule'#13#10;
          end;
          if not AnsiContainsText('ABCDEFGHIJKLMNOPQRSTUVWXYZ', TempCum) then
            LastErrorMsg := LastErrorMsg + 'Le deuxième caractère du cumul doit être une lettre majuscule ou un chiffre de 0 à 9'#13#10;
        end;
        if LastErrorMsg <> '' then
        begin
          PGIBox(LastErrorMsg,'Cumul');
          SetFocusControl('PCL_CUMULPAIE');
          exit;
        end;
      end;
      //FIN PT11

      if (not isnumeric(cumul)) then
      begin // Prise en compte du code alphanumérique
  {      TempCum := Copy(Cumul, 1, 1);
        if isnumeric(TempCum) then
        begin
          SetField('PCL_CUMULPAIE', '  ');
          PGIbox('Votre code cumul doit commencer par une lettre ', Ecran.Caption);
          SetFocusControl('PCL_CUMULPAIE');
          exit;
        end;}
        //DEB PT11
        if (DS.State = dsinsert) and (GetField('PCL_PREDEFINI') <> 'STD') then
        begin
          LastErrorMsg := 'La codification alphanumérique est réservée aux cumuls prédéfinis Standard'#13#10;
          PGIBox(LastErrorMsg,'Cumul');
          SetFocusControl('PCL_CUMULPAIE');
          exit;
          //SetField('PCL_PREDEFINI', 'STD');
          //SetControlEnabled('PCL_PREDEFINI', FALSE);
        //FIN PT11
        end;
      end
      else
      begin
        iCode := strtoint(trim(cumul));
        TempCum := ColleZeroDevant(iCode, 2);
        if (DS.State = dsinsert) and (TempCum <> '') and (pred <> '') then
        begin
          // PT2 10/04/2002 Ph V571 Accepttaion des cumuls pairs > 50 à STD
          OKRub := TestRubrique(pred, Cumul, 50, TRUE);
          if (OkRub = False) or (Cumul = '00') then
          begin
            Mes := MesTestRubrique('CUM', pred, 50);
            HShowMessage('2;Code Erroné: ' + Cumul + ' ;' + mes + ';W;O;O;;;', '', '');
            TempCum := '';
          end;
        end;
        if TempCum <> Cumul then
        begin
          SetField('PCL_CUMULPAIE', TempCum);
          SetFocusControl('PCL_CUMULPAIE');
        end;
      end;
    end;
  end;
  // FIN PT6
  if (F.FieldName = 'PCL_PREDEFINI') and (DS.State = dsinsert) then
  begin
    Cumul := (GetField('PCL_CUMULPAIE'));
    Pred := GetField('PCL_PREDEFINI');
    if Pred = '' then exit;

    AccesPredefini('TOUS', CEG, STD, DOS);
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      Error := True;
      PGIBox('Vous ne pouvez créer de cumul prédéfini CEGID', 'Accès refusé');
      Pred := 'DOS';
      SetControlProperty('PCL_PREDEFINI', 'Value', Pred);
    end;
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      Error := True;
      PGIBox('Vous ne pouvez créer de Cumul prédéfini Standard', 'Accès refusé');
      Pred := 'DOS';
      SetControlProperty('PCL_PREDEFINI', 'Value', Pred);
    end;
    if (Cumul <> '') and (Cumul[1] <> ' ') and (Pred <> '') and (IsNumeric (Cumul)) then // PT6 PT7
    begin
      // PT2 10/04/2002 Ph V571 Accepttaion des cumuls pairs > 50 à STD
      OKRub := TestRubrique(Pred, Cumul, 50, True);
      if (OkRub = False) or (Cumul = '00') then
      begin
        Mes := MesTestRubrique('CUM', Pred, 50);
        HShowMessage('2;Code Erroné: ' + Cumul + ' ;' + Mes + ';W;O;O;;;', '', '');
        SetField('PCL_CUMULPAIE', vide);
        if Pred <> GetField('PCL_PREDEFINI') then SetField('PCL_PREDEFINI', pred);
        SetFocusControl('PCL_CUMULPAIE');
        exit;
      end;
    end;
    if Pred <> GetField('PCL_PREDEFINI') then SetField('PCL_PREDEFINI', pred);
  end;

  if (ds.state in [dsBrowse]) then
  begin
    if LectureSeule then
    begin
      PaieLectureSeule(TFFicheListe(Ecran), True);
      SetControlEnabled('BDelete', False);
    end;
    SetControlEnabled('BInsert', True);
    SetControlEnabled('PCL_PREDEFINI', False);
    SetControlEnabled('PCL_RUBRIQUE', False);
  end;
end;

procedure TOM_CUMULPAIE.OnDeleteRecord;
var
  NomChamp: {array[1..1] of } string;
  ValChamp: {array[1..1] of } variant;
  ExisteCod: Boolean;
begin
  inherited;
  //DEB PT1
  { DEB PT4 Anomalie Test existence rubrique : utilisation du ##PREDEFINI## }
  NomChamp := '##PCR_PREDEFINI## PCR_CUMULPAIE';
  ValChamp := GetField('PCL_CUMULPAIE');
  //NomChamp[2]:='PCR_PREDEFINI'; ValChamp[2]:=GetField('PCL_PREDEFINI');
  //NomChamp[3]:='PCR_NODOSSIER'; ValChamp[3]:=GetField('PCL_NODOSSIER');
  ExisteCod := RechEnrAssocier('CUMULRUBRIQUE', [NomChamp], [Valchamp]);
  { FIN PT4 }
  if ExisteCod = TRUE then
  begin
    LastError := 1;
    LastErrorMsg := 'Attention! ce cumul est affecté à certaines rubriques. Suppression impossible!';
  end
  else
  begin
  //FIN PT1
// PT5  ChargementTablette(TFFicheListe(Ecran).TableName, '');
//PT9
    Trace := TStringList.Create ;  //PT9-2
    Trace.Add('SUPPRESSION CUMUL '+GetField('PCL_CUMULPAIE')+' '+ GetField('PCL_LIBELLE'));
    CreeJnalEvt('003','084','OK',nil,nil,Trace);
    FreeAndNil (Trace);  //PT9-2   Trace.free;
  end;
end;

procedure TOM_CUMULPAIE.OnLoadRecord;
begin
  inherited;
  AccesPredefini('TOUS', CEG, STD, DOS);
  LectureSeule := FALSE;
  if (Getfield('PCL_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFicheListe(Ecran), (CEG = False));
    //*  SetControlEnabled('BDefaire',CEG);
    SetControlEnabled('BDelete', CEG);
  end;
  if (Getfield('PCL_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFicheListe(Ecran), (STD = False));
    //*  SetControlEnabled('BDefaire',CEG);
    SetControlEnabled('BDelete', STD);
  end;
  if (Getfield('PCL_PREDEFINI') = 'DOS') then
  begin
    LectureSeule := (DOS = False);
    PaieLectureSeule(TFFicheListe(Ecran), LectureSeule);
    SetControlEnabled('BDelete', DOS);
  end;
  SetControlEnabled('BInsert', True);
  SetControlEnabled('PCL_PREDEFINI', False);
  SetControlEnabled('PCL_CUMULPAIE', False);

  if DS.State in [dsInsert] then
  begin
    LectureSeule := FALSE;
    PaieLectureSeule(TFFicheListe(Ecran), False);
    SetControlEnabled('PCL_PREDEFINI', TRUE);
    SetControlEnabled('PCL_CUMULPAIE', True);
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDelete', False);
    //*   SetControlEnabled('BDefaire',False);
  end
  else DerniereCreate := '';

end;


procedure TOM_CUMULPAIE.OnNewRecord;
begin
  inherited;
  SetField('PCL_PREDEFINI', 'DOS');
  SetField('PCL_THEMECUM', '001');
  SetField('PCL_TYPECUMUL', 'X');
  SetField('PCL_RAZCUMUL', '00');
end;

procedure TOM_CUMULPAIE.OnUpdateRecord;
var
  Predef, Cumul : string; // PT7
  TempCum,TempCum2 : String; //PT11
  LibErreur : String;
begin
  inherited;
  LastErrorMsg := '';
  Maj := True;
  Predef := GetField('PCL_PREDEFINI');
  LeStatut := DS.State; //PT12
  if (Predef <> 'CEG') and (Predef <> 'DOS') and (Predef <> 'STD') and (Error = False) then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez renseigner le champ prédéfini';
    SetFocusControl('PCL_PREDEFINI');
  end;
  if (DS.State = dsinsert) then
  begin
    DerniereCreate := GetField('PCL_CUMULPAIE');
    if (GetField('PCL_PREDEFINI') <> 'DOS') then
      SetField('PCL_NODOSSIER', '000000')
    else
      // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
      SetField('PCL_NODOSSIER', PgRendNoDossier());
  end;

  //DEB PT11 La lettre doit être une majuscule
  Cumul := GetField('PCL_CUMULPAIE');
  if (not isnumeric(Cumul)) then
  begin
    if (GetField('PCL_PREDEFINI') <> 'STD') then
    begin
      LastErrorMsg := 'La codification alphanumérique est réservée aux cumuls prédéfinis Standard'#13#10;
      LastError := 1;
      Error := True;
      SetFocusControl('PCL_CUMULPAIE');
    end
    else
    begin
      LibErreur := '';
      TempCum := Copy(Cumul, 1, 1);
      if not isnumeric(TempCum) then
      begin
        if (TempCum <> UpperCase(TempCum)) then
        begin
          TempCum := UpperCase(TempCum);
          TempCum2 := Copy(Cumul, 2, 1);
          Cumul := TempCum + TempCum2;
          LibErreur := 'La première lettre du cumul a été mis en majuscule'#13#10;
          SetField('PCL_CUMULPAIE', Cumul);
        end;
        if not AnsiContainsText('ABCDEFGHIJKLMNOPQRSTUVWXYZ', TempCum) then
          LibErreur := LibErreur + 'Le premier caractère du cumul doit être une lettre majuscule ou un chiffre de 0 à 9'#13#10;
      end;
      TempCum := Copy(Cumul, 2, 1);
      if not isnumeric(TempCum) and (GetField('PCL_PREDEFINI') = 'STD') then
      begin
        if (TempCum <> UpperCase(TempCum)) then
        begin
          TempCum := UpperCase(TempCum);
          TempCum2 := Copy(Cumul, 1, 1);
          Cumul := TempCum2 + TempCum;
          LibErreur := LibErreur + 'La deuxième lettre du cumul a été mis en majuscule'#13#10;
          SetField('PCL_CUMULPAIE', Cumul);
        end;
        if not AnsiContainsText('ABCDEFGHIJKLMNOPQRSTUVWXYZ', TempCum) then
          LibErreur := LibErreur + 'Le deuxième caractère du cumul doit être une lettre majuscule ou un chiffre de 0 à 9'#13#10;
      end;
      if LibErreur <> '' then
      begin
        LastErrorMsg := LibErreur;
        LastError := 1;
        Error := True;
      end;
    end;
  end;
  //FIN PT11

  // DEB PT7
  Cumul := Getfield('PCL_CUMULPAIE');
  if (Cumul = '') OR (Cumul = ' ') OR (Cumul = '  ') OR (length(Cumul) <> 2 )then //PT8
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez renseigner le code cumul avec 2 caractères';
    SetFocusControl('PCL_CUMULPAIE');
  end;
  // FIN PT7
  //Rechargement des tablettes
  if (LastError = 0) and (Getfield('PCL_CUMULPAIE') <> '') and (Getfield('PCL_LIBELLE') <> '') then
  begin
    // PT5    ChargementTablette(TFFicheListe(Ecran).TableName, '');
    SetControlEnabled('PCL_CUMULPAIE', False); //PT3 on force l'accès au champ ap. validation
  end;
  Error := False;
  Maj := False;
end;
//PT9
procedure TOM_CumulPaie.OnAfterUpdateRecord;
Var
    even: boolean;
    LaTable,lecode, LeLibelle : String;
begin
  LaTable := 'CUMULPAIE';
  LeCode := 'PCL_CUMULPAIE';
  LeLibelle := 'PCL_LIBELLE';
  Trace := TStringList.Create ; //PT9-2
  even := IsDifferent(dernierecreate,Latable,LeCode,LeLibelle,Trace,TFFicheListe(Ecran),LeStatut);  //PT37 //PT12
  FreeAndNil (Trace);  //PT9-2  Trace.free;

end;
procedure TOM_CUMULPAIE.OnArgument(stArgument: string);
begin
  inherited;
  PaieConceptPlanPaie(Ecran); // PT10
end;

initialization
  registerclasses([TOM_CUMULPAIE]);
end.

