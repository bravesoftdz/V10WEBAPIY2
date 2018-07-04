{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 01/03/2001
Modifié le ... : 02/10/2001
Description .. : TOM Affectation des codes DUCS aux rubriques de
               : cotisation.
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
{
 PT1 09/07/2001   V540 MF Ajout de contrôles avant validation
 PT2 02/10/2001   V552 MF Correction contrôle du PREDEFINI
 PT3 12/10/2001   V552 MF Correction : Contrôle que la codification saisie existe.
 PT4 13/05/2002   V575 MF
                  1-Correction fiche qualité 10083. Après une suppression
                    la fiche est effacée et on revient sur la liste à jour
                  2-Correction fiche qualité 10107. On ne permet pas d'affecter
                    une rubrique à une codification de type "Taux AT".
 PT5 24/10/2002   V585 MF
                  1-L'affectation en STD ou CEG d'une rubrique de type DOS n'est
                  pas possible. De même, l'affectation CEG d'une rubrique de
                  type DOS ou STD n'est pas possible .
                  2-Le bouton DUPLIQUER n'est plus visible
                  3-Message d'alerte quand en création il existe déja une
                  même affectation pour un autre niveau de prédéfini.
 PT6 17/03/2003   V 420 MF
                  1- Fiche qualité n° 10531 : utilisation ##PREDEFINI## pour
                  ne séléctionner que les codification concernant le dossier
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****

 PT7 09/02/2007   V 700 MF
                  FQ 13070 : Traitement des codifications ALSACE MOSELLE

 PT8 12/02/2007   V8.00 GGS FQ 12694 Journal d'événements
 PT8-2 26/04/2007 V8.00 GGS Revue gestion de Trace
 PT9 28/06/2007   V7_02 MF  FQ 14488 : Accès à la fiche Affectation Ducs en création et modification
 PT10 24/07/2007  V7_02 MF  FQ 14603 : Filtre de la liste des codification en fct de la nature de l'organisme
 PT11 FC 06/11/2007 V_80 Correction bug journal des évènement quand création (lib Modification au lieu de Création)
}

unit UTOMDUCSAFFECT;

interface

uses
//unused Windows,
//unused  StdCtrls,
  Controls, Classes,
//unused  Graphics,
  forms, sysutils,
//unused  ComCtrls,
//unused  HTB97,
  {$IFDEF EAGLCLIENT}
//unused  UtileAGL,
  eFiche,
//  UTOB, //unused
  {$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}    //PT7 (compile)
  HDB,
//unused  DBCtrls,
  Fiche,
  {$ENDIF}
//unused  Grids,
  HCtrls,
//unused  HEnt1,
//unused  EntPaie,
  HMsgBox, UTOM,
//unused  UTOB,
//unused  Vierge,
  P5Def,                //pt8 enlevé le commentaire unused
//unused  AGLInit,
  PGOutils,PgOutils2,
  UTOB, 
  PAIETOM;              //PT8

type
  TOM_DUCSAFFECT = class(PGTOM)   //PT8 TOM devient PGTOM
  private
    PGPredef, PGNodossier, PGCotisat: string;
    CEG, STD, DOS, LectureSeule: Boolean;
    Loaded: boolean;
    MsgError: integer;
    Trace: TStringList;                   //PT8
    DerniereCreate: string;               //PT8
    OnFerme: Boolean;                     //PT8
    LeStatut:TDataSetState; //PT11
    function filtrecodif(Rubrique : string) : string; // PT10

public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoadRecord; override;
    procedure OnNewRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnUpdateRecord; override;
    procedure OnDeleteRecord; override; // PT4-1
    procedure OnAfterUpdateRecord; override;    //PT8
  end;

implementation


procedure TOM_DUCSAFFECT.OnArgument(Arguments: string);
var
  st: string;
//  Q                 : TQuery; // PT9

begin
  inherited;

  SetControlProperty('BDUPLIQUER', 'VISIBLE', FALSE); // PT5-2

  st := Trim(Arguments);

  PGPredef := '';
  PGNodossier := '';
  PGCotisat := '';
//PT8-2  Trace := TStringList.Create ;           //PT8
  if St = '' then exit;     // on vient du multi critères des cotis. affectées

  PGPredef := ReadTokenSt(st); // Recup Predefini
  if Copy(PGPredef, 1, 3) = 'ACT' then
  begin // test si argument = Action=creation
    PGPredef := '';
    exit;
  end;
  PGNodossier := ReadTokenSt(st); // Recup Nodossier
  PGCotisat := ReadTokenSt(st); // Recup Cotisation
  AccesPredefini('TOUS', CEG, STD, DOS);  // PT2
end;


procedure TOM_DUCSAFFECT.OnChangeField(F: TField);
var
  Pred, Rubrique    : string;
  LpredAf, LpredRu  : string; //PT5
  Q                 : TQuery; //PT7

begin
  inherited;
  MsgError := 0;
  if Loaded then
  begin
    Loaded := FALSE;
    exit;
  end;
  if (F.FieldName = 'PDF_PREDEFINI') and (DS.State = dsinsert) then
  begin
    Pred := GetField('PDF_PREDEFINI');
    Rubrique := (GetField('PDF_RUBRIQUE'));
    if Pred = '' then exit;
    AccesPredefini('TOUS', CEG, STD, DOS);
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez créer d''affectation prédéfini CEGID', 'Accès refusé'); //PT2
      SetControlProperty('PDF_PREDEFINI', 'Value', 'DOS');
      SetField('PDF_PREDEFINI', 'DOS');
      MsgError := 1;
    end;
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez créer d''affectation prédéfini Standard', 'Accès refusé'); //PT2
      SetControlProperty('PDF_PREDEFINI', 'Value', 'DOS');
      SetField('PDF_PREDEFINI', 'DOS');
      MsgError := 1;
    end;
    // d PT5
    if (MsgError = 0) then
    begin
      if ((PGPredef = 'DOS') and (Pred <> 'DOS') or
        (PGPredef = 'STD') and (Pred = 'CEG')) then
      begin
        LPredAf := '';
        LPredRu := '';
        if (PGPredef = 'DOS') then LPredRu := 'Dossier';
        if (PGPredef = 'STD') then LPredRu := 'Standard';

        if (Pred = 'STD') then LPredAf := 'Standard';
        if (Pred = 'CEG') then LPredAf := 'CEGID';
        PgiBox('Il n''est pas possible de céer une affectation "' + LPredAf +
          '" pour une rubrique de type "' + LPredRu + '"', 'Accès refusé');
        SetControlProperty('PDF_PREDEFINI', 'Value', PGPredef);
        SetField('PDF_PREDEFINI', PGPredef);
        MsgError := 1;
      end;
    end;
    // f PT5
    PGPredef := GetField('PDF_PREDEFINI');
  end;
  if (ds.state in [dsBrowse]) then
  begin
    if LectureSeule then
    begin
      PaieLectureSeule(TFFiche(Ecran), True);
      SetControlEnabled('BDelete', False);
    end;
    SetControlEnabled('BInsert', True);
    SetControlEnabled('PDF_PREDEFINI', False);
    SetControlEnabled('PDF_RUBRIQUE', False);
  end;
// d PT7
  if (F.FieldName = 'PDF_CODIFICATION') and (DS.State = dsinsert) then
  begin
    SetControlText ('PDF_CODIFALSACE','');
    if (Getfield('PDF_CODIFICATION') <> '') then
    begin
      Q := OpenSql('SELECT PDP_CODIFALSACE' +
                   ' FROM DUCSPARAM ' +
                   ' WHERE ##PDP_PREDEFINI## AND PDP_CODIFICATION = "' +
                   Getfield('PDF_CODIFICATION')  + '"', True);
    if not Q.eof then
    begin
      SetControlText ('PDF_CODIFALSACE',Q.FindField('PDP_CODIFALSACE').AsString);
      SetField ('PDF_CODIFALSACE',Q.FindField('PDP_CODIFALSACE').AsString);   // PT9
    end;
    Ferme(Q);
  end;

  if ((Getfield('PDF_CODIFICATION') >='1000000') and
      (Getfield('PDF_CODIFICATION') <='1ZZZZZZ')) then
  begin
    SetControlVisible('PDF_CODIFALSACE', True);
    SetControlVisible('TPDF_CODIFALSACE', True);
  end
  else
  begin
    SetControlVisible('PDF_CODIFALSACE', False);
    SetControlVisible('TPDF_CODIFALSACE', False);
  end;
  end;
// f PT7
end;

procedure TOM_DUCSAFFECT.OnLoadRecord;
// d PT10
var
{$IFNDEF EAGLCLIENT}
  CodifAssocie                  : THDBEdit;
{$ELSE}
  CodifAssocie                  : THEdit;
{$ENDIF}
// f PT10
begin
  inherited;
  Loaded := True;
  if not (DS.State in [dsInsert]) then DerniereCreate := '';

  AccesPredefini('TOUS', CEG, STD, DOS);
  LectureSeule := FALSE;
  if (Getfield('PDF_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
    SetControlEnabled('BDelete', CEG);
  end;
  if (Getfield('PDF_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('BDelete', STD);
  end;
  if (Getfield('PDF_PREDEFINI') = 'DOS') then
  begin
    LectureSeule := False;
    PaieLectureSeule(TFFiche(Ecran), (DOS = False));
    SetControlEnabled('BDelete', DOS);
  end;

  SetControlEnabled('BInsert', True);
  SetControlEnabled('PDF_PREDEFINI', False);
  SetControlEnabled('PDF_RUBRIQUE', False);

  if DS.State in [dsInsert] then
  begin
    LectureSeule := FALSE;
    PaieLectureSeule(TFFiche(Ecran), False);
    SetControlEnabled('PDF_PREDEFINI', True);
    SetControlEnabled('PDF_RUBRIQUE', True);
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDelete', False);
  end;
// d PT7
  SetControlEnabled('PDF_CODIFALSACE', False);
  if ((Getfield('PDF_CODIFICATION') >='1000000') and
      (Getfield('PDF_CODIFICATION') <='1ZZZZZZ')) then
  begin
    SetControlVisible('PDF_CODIFALSACE', True);
    SetControlVisible('TPDF_CODIFALSACE', True);
    setControlText('PDF_CODIFALSACE',GetField('PDF_CODIFALSACE')); // PT9
  end
  else
  begin
    SetControlVisible('PDF_CODIFALSACE', False);
    SetControlVisible('TPDF_CODIFALSACE', False);
  end;
  // f PT7
    LectureSeule := FALSE;
  if (Getfield('PDF_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
    SetControlEnabled('BDelete', CEG);
  end;
  if (Getfield('PDF_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('BDelete', STD);
  end;
  if (Getfield('PDF_PREDEFINI') = 'DOS') then
  begin
    LectureSeule := False;
    PaieLectureSeule(TFFiche(Ecran), (DOS = False));
    SetControlEnabled('BDelete', DOS);
  end;
  PaieConceptPlanPaie(Ecran);
// d PT10
{$IFNDEF EAGLCLIENT}
         CodifAssocie := THDBEdit(GetControl('PDF_CODIFICATION'));
{$ELSE}
         CodifAssocie := THEdit(GetControl('PDF_CODIFICATION'));
{$ENDIF}
//select distinct(POG_NATUREORG), POG_ORGANISME,PCT_RUBRIQUE from organismepaie left join cotisation on pct_organisme=pog_organisme where ##POG_PREDEFINI## and PCT_RUBRIQUE="8096"
  CodifAssocie.plus := filtrecodif(GetField('PDF_RUBRIQUE'));
// f PT10

end;

procedure TOM_DUCSAFFECT.OnNewRecord;
begin
  inherited;
  Loaded := TRUE;
  if PGPredef <> '' then SetField('PDF_PREDEFINI', PGPredef);
  if (PGPredef = 'CEG') and (CEG = FALSE) then SetField('PDF_PREDEFINI', 'DOS'); // PT2
  if PGNodossier <> '' then SetField('PDF_NODOSSIER', PGNodossier);
  if PGCotisat <> '' then SetField('PDF_RUBRIQUE', PGCotisat);
end;

{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 01/03/2001
Modifié le ... : 09/07/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DUCSAFFECT.OnUpdateRecord;
var
  st: string;
  //PT4-2   Champ :array[1..1] of string;
  //PT4-2   Valeur :array[1..1] of variant ;
  Champ: array[1..2] of Hstring; //PT4-2
  Valeur: array[1..2] of variant; //PT4-2
  ExisteCod: Boolean;
begin
  inherited;
  OnFerme := False;                                           //PT8
  LeStatut := DS.State; //PT11
  if (DS.State = dsinsert) then
  begin
    DerniereCreate := GetField('PDF_RUBRIQUE');              //PT8
    if (GetField('PDF_PREDEFINI') <> 'DOS') then
      SetField('PDF_NODOSSIER', '000000')
    else
      // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
      SetField('PDF_NODOSSIER', PgRendNoDossier());
      if (DerniereCreate = GetField('PDF_RUBRIQUE')) then OnFerme := True; //PT8 le bug arrive on se casse !!!
  end;
  // début PT1
  st := '';
  st := GetField('PDF_CODIFICATION');
  if st = '' then
  begin
    MsgError := 1;
    SetFocusControl('PDF_CODIFICATION');
    LastErrorMsg := 'La codification est obligatoirement renseignée';
  end;
  st := '';
  st := GetField('PDF_RUBRIQUE');
  if st = '' then
  begin
    MsgError := 1;
    SetFocusControl('PDF_RUBRIQUE');
    LastErrorMsg := 'Le code rubrique de cotisation est obligatoirement renseigné';
  end;
  st := '';
  st := GetField('PDF_PREDEFINI');
  if st = '' then
  begin
    MsgError := 1;
    SetFocusControl('PDF_PREDEFINI');
    LastErrorMsg := 'Le prédéfini est obligatoirement renseigné';
  end;
  st := '';
  // fin PT1
  // PT3 début
  //PT6   Champ[1]:= 'PDP_CODIFICATION';
  Champ[1] := '##PDP_PREDEFINI## PDP_CODIFICATION';
  Valeur[1] := GetField('PDF_CODIFICATION');
  Champ[2] := '';
  Valeur[2] := ''; //PT4-2

  ExisteCod := RechEnrAssocier('DUCSPARAM', Champ, Valeur);
  if (ExisteCod = FALSE) then
  begin
    MsgError := 1;
    SetFocusControl('PDF_CODIFICATION');
    LastErrorMsg := 'Cette codification n''existe pas';
  end;
  // PT3 fin

  // PT4-2 début
  //PT6   Champ[1]:= 'PDP_CODIFICATION';
  Champ[1] := '##PDP_PREDEFINI## PDP_CODIFICATION';
  Valeur[1] := GetField('PDF_CODIFICATION');
  Champ[2] := 'PDP_TYPECOTISATION';
  Valeur[2] := 'A'; //PT4-2

  ExisteCod := RechEnrAssocier('DUCSPARAM', Champ, Valeur);
  if (ExisteCod = TRUE) then
  begin
    MsgError := 1;
    SetFocusControl('PDF_CODIFICATION');
    LastErrorMsg := 'Cette codification est reservée et ne peut pas être affectée';
  end;

  // PT4-2 fin
  // d PT5-3
  Champ[1] := 'PDF_CODIFICATION';
  Valeur[1] := GetField('PDF_CODIFICATION');
  //PT6   Champ[2]:= 'PDF_RUBRIQUE';
  Champ[2] := '##PDF_PREDEFINI## PDF_RUBRIQUE';
  Valeur[2] := GetField('PDF_RUBRIQUE');
  ExisteCod := RechEnrAssocier('DUCSAFFECT', Champ, Valeur);
  if (ExisteCod = TRUE) then
  begin
    if PGIAsk('Cette affectation existe déjà à un autre niveau de prédéfini.#13#10' +
      ' Confirmez-vous la création?', 'Affectation') <> mrYes then
    begin
      MsgError := 1;
      SetFocusControl('PDF_CODIFICATION');
    end;
  end;
  // f PT5-3

  if (LastError <> 0) then SetField('PDF_PREDEFINI', GetField('PDF_PREDEFINI')); // MF
  LastError := MsgError;
end;
// début PT4-1
procedure TOM_DUCSAFFECT.OnDeleteRecord;
begin
  inherited;
//PT8
    Trace := TStringList.Create ; //PT8-2
    Trace.Add('SUPPRESSION AFFECTATION DUCS '+GetField('PDF_RUBRIQUE')+' '+ GetField('PDF_CODIFICATION'));
    CreeJnalEvt('003','091','OK',nil,nil,Trace);
    FreeAndNil (Trace);         //PT8-2 Trace.free;
  Ecran.Close;
end;
// fin PT4-1

{procedure TOM_DUCSAFFECT.DupliquerAffectation(Sender: TObject);
var
   Ok : Boolean;
   Champ :array[1..3] of string;
   Valeur :array[1..3] of variant ;
   NoDossier, Champlib, ChamplibS, ChampArrB, ChampArrM, ChampTyp : string;}
{$IFNDEF EAGLCLIENT}
{    Code : THDBEdit;}
{$ELSE}
{    Code : THEdit;}
{$ENDIF}

{begin}
{ TFFiche(Ecran).BValider.Click;
 mode:='DUPLICATION';
 AglLanceFiche('PAY','CODE', '', '', 'AFF; ; ;7');
 if PGCodeDupliquer<>'' then
   begin
   Champ[1]:= 'PDP_PREDEFINI';  Valeur[1]:= PGCodePredefini;
   Champ[2]:= 'PDP_NODOSSIER';
   if PGCodePredefini='DOS' then Valeur[2]:= V_PGI_env.nodossier
   else                          Valeur[2]:= '000000';
   Champ[3]:= 'PDP_CODIFICATION';   Valeur[3]:= PGCodeDupliquer;
   Ok:=RechEnrAssocier('DUCSPARAM',Champ,Valeur);
   if Ok=False then         //Test si code existe ou non
     begin     }
{$IFNDEF EAGLCLIENT}
{      Code:=THDBEdit(GetControl('PDP_CODIFICATION'));}
{$ELSE}
{      Code:=THEdit(GetControl('PDP_CODIFICATION'));}
{$ENDIF}
{      if (code<>nil) then DupliquerPaie (TFFiche(Ecran).TableName,Ecran);
      SetField('PDP_CODIFICATION',PGCodeDupliquer);
      SetField('PDP_PREDEFINI',PGCodePredefini);
      AccesFicheDupliquer(TFFiche(Ecran),PGCodePredefini,NoDossier,LectureSeule);
      SetField('PDP_NODOSSIER',NoDossier);
      SetControlEnabled('PDP_CODIFICATION',False);
      SetControlEnabled('PDP_PREDEFINI',False);
      SetField('PDP_LIBELLE',ChampLib);
      SetField('PDP_LIBELLESUITE',ChampLibS);
      SetField('PDP_BASETYPARR',ChampArrB);
      SetField('PDP_MTTYPARR',ChampArrM);
      SetField('PDP_TYPECOTISATION',ChampTyp);
      TFFiche(Ecran).Bouge(nbPost) ; //Force enregistrement
     end
   else
     HShowMessage ('5;Codification :;La duplication est impossible, la codification existe déjà.;W;O;O;O;;;','','');

   end;
end; }
procedure TOM_DUCSAFFECT.OnAfterUpdateRecord;  //PT8
var
    even: boolean;

begin
  inherited;
  Trace := TStringList.Create ;   //PT8-2
  even := IsDifferent(dernierecreate,PrefixeToTable(TFFiche(Ecran).TableName),'PDF_RUBRIQUE',TFFiche(Ecran).LibelleName,Trace,TFFiche(Ecran),LeStatut); //PT11
  FreeAndNil (Trace);   // PT8-2 Trace.free;
  if OnFerme then Ecran.Close;
end;

// d PT10
function TOM_DUCSAFFECT.filtrecodif(Rubrique : string) : string;
var
  st                : string;
  Q                 : TQuery;
  Tob_NatureOrg, T  : TOB;

begin
  Result := '';


  st := 'select distinct(POG_NATUREORG), POG_ORGANISME,PCT_RUBRIQUE '+
        'from organismepaie left join cotisation on pct_organisme=pog_organisme'+
        ' where ##POG_PREDEFINI## and PCT_RUBRIQUE="'+Rubrique+'"';
  Q := OpensQl(st, True);

  Tob_NatureOrg := TOB.Create('Les Nature Organisme', nil, -1);
  Tob_NatureOrg.loaddetaildb('NATUREORG', '', '', Q, false);
  if Tob_NatureOrg.detail.count = 1 then
  begin
    T := Tob_NatureOrg.Detail[0];
    st := T.GetValue('POG_NATUREORG');
    end
  else
    st := '';

  Ferme(Q);
  FreeAndNil(Tob_NatureOrg);

  if (st <> '') then
  begin
    if (st = '100') then
    // URSSAF
      Result := ' AND PDP_CODIFICATION >= "1" AND PDP_CODIFICATION <="1ZZZZZZ"'
    else
    if (st = '200') then
    // ASSEDIC
      Result := ' AND PDP_CODIFICATION >= "2" AND PDP_CODIFICATION <="2ZZZZZZ"'
    else
    if (st = '300') then
    // IRC
       Result := ' AND PDP_CODIFICATION >= "3" AND PDP_CODIFICATION <="5ZZZZZZ"'+
                 ' OR PDP_CODIFICATION >="8" AND PDP_CODIFICATION <="9ZZZZZZ"'
    else
    if (st = '600') then
    // MSA
      Result := ' AND PDP_CODIFICATION >= "6" AND PDP_CODIFICATION <="6ZZZZZZ"'
    else
    if (st = '700') then
    // BTP
      Result := ' AND PDP_CODIFICATION >= "7" AND PDP_CODIFICATION <="7ZZZZZZ"';
  end;
{$IFDEF EAGLCLIENT}
  if (st = '') then
    Result := ' AND PDP_CODIFICATION >= "1" AND PDP_CODIFICATION <= "9ZZZZZZ"';
{$ENDIF}
end;
// f PT10

initialization
  registerclasses([TOM_DUCSAFFECT]);
end.

