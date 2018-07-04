{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 16/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGTABLESDYNADUP ()
Mots clefs ... : TOF;PGTABLESDYNADUP
*****************************************************************}
{
PT1 31/07/2007 V8 GGU gestion du niveau de saisie pour toutes les natures de table 
PT2 04/10/2007 FC V_80 Concept plan de paie
PT3 21/04/2008 V81 GGU FQ 15364 Duplication multi-dossiers
}
Unit UTofPG_TABLESDYNADUP;

Interface

Uses StdCtrls, Controls, Classes,
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, uTob,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, HTB97 ;

Type
  TOF_PGTABLESDYNADUP = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    BtnCancel, BtnValider : TToolBarButton97;
    Control_VALNIV, Control_PREDEFINI, Control_NIVSAIS : THValComboBox;
    Control_DTVALID : THEdit;
    CEG, STD, DOS : Boolean;  //, Autorisation
    PTE_Key, PTD_Key, Action, CODTABL, PREDEFINI, PREDEFINIinit, NODOSSIER, DTVALID, NIVSAIS, VALNIV, NATURETABLE : String;
    ValCritere1, ValCritere2 : String;
    procedure OnChangePredefini(Sender : TObject);
    procedure OnChangeNIVSAIS(Sender : TObject);
    procedure OnChangeDTVALID(Sender : TObject);
    procedure OnChangeVALNIV(Sender : TObject);
    procedure OnValidate(Sender : TObject);
  end ;

Implementation

uses
  PGOUTILS, ed_tools, ComObj
  ;

procedure TOF_PGTABLESDYNADUP.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGTABLESDYNADUP.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGTABLESDYNADUP.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGTABLESDYNADUP.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGTABLESDYNADUP.OnArgument (S : String ) ;
var
  Arg : String;
  Q : TQuery;
  stQ : String;
  procedure MajKey;
  begin
    PTE_Key := '( PTE_CODTABL   ="'+CODTABL   +'"'
          +'  AND PTE_NODOSSIER ="'+NODOSSIER +'"'
          +'  AND PTE_DTVALID   ="'+UsDateTime(AGLStrToDate(DTVALID))+'"'
          +'  AND PTE_NIVSAIS   ="'+NIVSAIS   +'"'
          +'  AND PTE_VALNIV    ="'+VALNIV    +'"'
          +'  AND PTE_PREDEFINI ="'+PREDEFINI +'")';
    PTD_Key := '( PTD_CODTABL   ="'+CODTABL   +'"'
          +'  AND PTD_NODOSSIER ="'+NODOSSIER +'"'
          +'  AND PTD_DTVALID   ="'+UsDateTime(AGLStrToDate(DTVALID))+'"'
          +'  AND PTD_NIVSAIS   ="'+NIVSAIS   +'"'
          +'  AND PTD_VALNIV    ="'+VALNIV    +'"'
          +'  AND PTD_PREDEFINI ="'+PREDEFINI +'")';
  end;
begin
  Inherited ;
  Arg := S;
  Action := Trim(ReadTokenPipe(Arg, ';')); //On récupère le type d'action
  if (Action = 'ACTION=DUPLICATION') then Action := 'DUPLICATION';
  CODTABL   := Trim(ReadTokenPipe(Arg, ';'));
  PREDEFINI := Trim(ReadTokenPipe(Arg, ';'));
  NODOSSIER := Trim(ReadTokenPipe(Arg, ';'));
  DTVALID   := Trim(ReadTokenPipe(Arg, ';'));
  NIVSAIS   := Trim(ReadTokenPipe(Arg, ';'));
  VALNIV    := Trim(ReadTokenPipe(Arg, ';'));
  NATURETABLE:= Trim(ReadTokenPipe(Arg, ';'));
  stQ   := 'SELECT PTE_CRITERE1, PTE_CRITERE2, PTE_LIBELLE '
         + ' FROM TABLEDIMENT '
         + 'WHERE PTE_CODTABL   ="'+CODTABL+'"'
         + '  AND PTE_PREDEFINI ="'+PREDEFINI+'"'
         + '  AND PTE_NODOSSIER ="'+NODOSSIER+'"'
         + '  AND PTE_DTVALID   ="'+UsDateTime(AGLStrToDate(DTVALID))+'"'
         + '  AND PTE_NIVSAIS   ="'+NIVSAIS+'"'
         + '  AND PTE_VALNIV    ="'+VALNIV +'"';
  Q := OpenSQL(stQ,True,1);
  if not Q.Eof then
  begin
    ValCritere1 := Q.FindField('PTE_CRITERE1').AsString;
    ValCritere2 := Q.FindField('PTE_CRITERE2').AsString;
    Ecran.Caption := 'Duplication de la table dynamique '+CODTABL+' '+Q.FindField('PTE_LIBELLE').AsString;
  end;
  Ferme(Q);

  BtnCancel := (GetControl('BFERME')    as TToolBarButton97);
  BtnValider := (GetControl('BVALIDER') as TToolBarButton97);
  Control_PREDEFINI:= (GetControl('_PREDEFINI') as THValComboBox);
  Control_NIVSAIS  := (GetControl('_NIVSAIS')   as THValComboBox);
  Control_VALNIV   := (GetControl('_VALNIV')    as THValComboBox);
  Control_DTVALID   := (GetControl('_DTVALID')  as THEdit);

//  On initialise les champs
  SetControlText('_NIVSAIS',NIVSAIS);
  if VALNIV <> '...' then SetControlText('_VALNIV',VALNIV) else SetControlText('_VALNIV','');
  SetControlText('_DTVALID',DTVALID);
  AccesPredefini('TOUS', CEG, STD, DOS);
//  Autorisation := True;
  SetControlText('_PREDEFINI',PREDEFINI);

  //PT2
  if (CEG = False) and (STD = False) then
    SetControlProperty('_PREDEFINI','Plus',' AND CO_CODE="DOS"');

  PREDEFINIinit := PREDEFINI;

  OnChangePredefini(self);

//  if (PREDEFINI = 'CEG') and (CEG = FALSE) then
//  begin
//    Autorisation := False;
//    PGIBox(TraduireMemoire('Vous n''avez pas l''autorisation de dupliquer les tables prédéfini CEGID'));
//  end else if (PREDEFINI = 'STD') and (STD = FALSE) then
//  begin
//    Autorisation := False;
//    PGIBox(TraduireMemoire('Vous n''avez pas l''autorisation de dupliquer les tables prédéfini STANDARD'));
//  end;
  setControlProperty('_NIVSAIS','Plus','and co_libre like "%'+GetControlText('_PREDEFINI')+'%"');
  SetControlEnabled('LBL_NIVSAIS', True); //PT1 (NATURETABLE <> 'COD')
  SetControlEnabled('_NIVSAIS',    True); //PT1 (NATURETABLE <> 'COD')
  SetControlEnabled('_VALNIV',     True); //PT1 (NATURETABLE <> 'COD')
//  Control_PREDEFINI.Enabled :=    ((ValCritere1 = '1500') or (ValCritere1 = '1501') or (ValCritere1 = ''))
//                              and ((ValCritere2 = '1500') or (ValCritere2 = '1501') or (ValCritere2 = ''));
  Control_PREDEFINI.Enabled := (PREDEFINI <> 'DOS');
  Control_PREDEFINI.OnChange := OnChangePREDEFINI;
  Control_NIVSAIS.OnChange   := OnChangeNIVSAIS;
  Control_VALNIV.OnChange    := OnChangeVALNIV;
  Control_DTVALID.OnChange   := OnChangeDTVALID;
  OnChangeNIVSAIS(nil);
  BtnValider.onClick := OnValidate;
  MajKey;
end ;

procedure TOF_PGTABLESDYNADUP.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGTABLESDYNADUP.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGTABLESDYNADUP.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PGTABLESDYNADUP.OnChangePredefini(Sender: TObject);
begin
  BtnValider.Enabled := True; //Autorisation;
  //PT2
  if (GetControlText('_PREDEFINI') = '') then
    SetControlText('_PREDEFINI','DOS');
  if (GetControlText('_PREDEFINI') = 'CEG') and (CEG = FALSE) then
  begin
    PGIBox(TraduireMemoire('Vous ne pouvez pas créer de prédéfini CEGID.'), Ecran.Caption);
    SetControlText('_PREDEFINI','DOS');
  end else if (GetControlText('_PREDEFINI') = 'STD') and (STD = FALSE) then
  begin
    PGIBox(TraduireMemoire('Vous ne pouvez pas créer de prédéfini Standard.'), Ecran.Caption);
    SetControlText('_PREDEFINI','DOS');
  end else if   ((PREDEFINIinit = 'STD') and (GetControlText('_PREDEFINI') =  'CEG'))
             or ((PREDEFINIinit = 'DOS') and (GetControlText('_PREDEFINI') <> 'DOS')) then
  begin
    PGIBox(TraduireMemoire('Vous ne pouvez pas dupliquer une table vers un prédéfini de niveau supérieur au niveau initial.'), Ecran.Caption);
    SetControlText('_PREDEFINI','DOS');
  end;
  setControlProperty('_NIVSAIS','Plus','and co_libre like "%'+GetControlText('_PREDEFINI')+'%"');
end;

procedure TOF_PGTABLESDYNADUP.OnValidate(Sender: TObject);
var
  stSQLInsert : String;
  PrioriteTemp : Integer;
  ValueNoDossier : String;
begin
  {//PT3 Si on duplique sur du dossier, on met V_PGI.NoDossier dans le nodossier,
    sinon, on met 000000 }
  ValueNoDossier := NODOSSIER; //PT3
  if GetControlText('_PREDEFINI') <> PREDEFINI then  //PT3
  begin
    if GetControlText('_PREDEFINI') = 'DOS' then
      ValueNoDossier := V_PGI.NoDossier
    else
      ValueNoDossier := '000000';
  end;
  if GetControlText('_NIVSAIS') = '' then
  begin
    LastErrorMsg := TraduireMemoire('Vous devez choisir le niveau de la saisie.');
    SetFocusControl('_NIVSAIS');
    LastError := 101;
    PGIBox(LastErrorMsg, TraduireMemoire('Duplication de tables dynamiques'));
    Exit;
  end;
  if (GetControlText('_DTVALID') = '') or (AGLStrToDate(GetControlText('_DTVALID')) = iDate1900) then
  begin
    LastErrorMsg := TraduireMemoire('Vous devez choisir une date de validité.');
    SetFocusControl('_DTVALID');
    LastError := 103;
    PGIBox(LastErrorMsg, TraduireMemoire('Duplication de tables dynamiques'));
    Exit;
  end;
  //Les données sont bonnes, on duplique
  PrioriteTemp := 0;
  if      GetControlText('_PREDEFINI') = 'CEG' then PrioriteTemp := 1000
  else if GetControlText('_PREDEFINI') = 'STD' then PrioriteTemp := 2000
  else if GetControlText('_PREDEFINI') = 'DOS' then PrioriteTemp := 3000;
  if      GetControlText('_NIVSAIS') = 'GEN' then PrioriteTemp := PrioriteTemp + 100
  else if GetControlText('_NIVSAIS') = 'CON' then PrioriteTemp := PrioriteTemp + 200
  else if GetControlText('_NIVSAIS') = 'ETB' then PrioriteTemp := PrioriteTemp + 300;
  if (GetControlText('_NIVSAIS') = 'CON') and (VALNIV <> '000') then PrioriteTemp := PrioriteTemp + 20;
  if (GetControlText('_NIVSAIS') = 'ETB') and (VALNIV <> '...') then PrioriteTemp := PrioriteTemp + 30;
  try
    stSQLInsert := 'INSERT INTO TABLEDIMENT (PTE_CODTABL, PTE_PREDEFINI, '
                 + 'PTE_NODOSSIER, PTE_DTVALID, PTE_NIVSAIS, PTE_VALNIV, PTE_PRIORITENIV, PTE_LIBELLE, '
                 + 'PTE_NATURETABLE, PTE_CRITERE1, PTE_CRITERE2, PTE_TYPECRITERE1, '
                 + 'PTE_TYPECRITERE2, PTE_SENS1, PTE_SENS2, PTE_SENSINT, PTE_TYPERESULTAT, '
                 + 'PTE_DATECREATION, PTE_DATEMODIF, PTE_CREATEUR, '
                 + 'PTE_UTILISATEUR) '
                 + 'SELECT "'+CODTABL+'", "'+GetControlText('_PREDEFINI')+'", "'
                 + ValueNoDossier+'", "'+USDATETIME(AGLStrToDate(GetControlText('_DTVALID')))+'", "'
                 + GetControlText('_NIVSAIS')+'", "'+VALNIV+'", '+intToStr(PrioriteTemp)+', PTE_LIBELLE , '
                 + 'PTE_NATURETABLE, PTE_CRITERE1, PTE_CRITERE2, PTE_TYPECRITERE1, '
                 + 'PTE_TYPECRITERE2, PTE_SENS1, PTE_SENS2, PTE_SENSINT, PTE_TYPERESULTAT, '
                 + 'PTE_DATECREATION, PTE_DATEMODIF, PTE_CREATEUR, '
                 + 'PTE_UTILISATEUR FROM TABLEDIMENT '
                 + 'WHERE '+PTE_Key;
    ExecuteSQL ( stSQLInsert );
    stSQLInsert := 'INSERT INTO TABLEDIMDET (PTD_CODTABL, PTD_PREDEFINI, '
                 + 'PTD_NODOSSIER, PTD_DTVALID, PTD_NIVSAIS, PTD_VALNIV, PTD_VALCRIT1, '
                 + 'PTD_VALCRIT2, PTD_LIBELLECODE, PTD_RESULTAT, PTD_DATECREATION, '
                 + 'PTD_DATEMODIF, PTD_CREATEUR, PTD_UTILISATEUR) '
                 + 'SELECT "'+CODTABL+'", "'+GetControlText('_PREDEFINI')+'", "'
                 + ValueNoDossier+'", "'+USDATETIME(AGLStrToDate(GetControlText('_DTVALID')))+'", "'
                 + GetControlText('_NIVSAIS')+'", "'+VALNIV+'", PTD_VALCRIT1, '
                 + 'PTD_VALCRIT2, PTD_LIBELLECODE, PTD_RESULTAT, PTD_DATECREATION, '
                 + 'PTD_DATEMODIF, PTD_CREATEUR, PTD_UTILISATEUR FROM TABLEDIMDET '
                 + 'WHERE '+PTD_Key;
    ExecuteSQL ( stSQLInsert );
  except
    on E : EOleException do
    begin
      if Pos('duplicate key', E.Message) > 0 then
        PGIError('Cette table existe déjà.', Ecran.Caption)
      else
        PGIError('La duplication a échouée.#10#13'+E.Message, Ecran.Caption);
    end;
    on E : Exception do
      PGIError('La duplication a échouée.#10#13'+E.Message, Ecran.Caption);
  end;
end;

procedure TOF_PGTABLESDYNADUP.OnChangeNIVSAIS(Sender: TObject);
var
  st : string;
begin
  st := GetControlText('_NIVSAIS');
  if st = 'ETB' then
  begin
    setControlProperty('_VALNIV','Vide',True);
    setControlProperty('_VALNIV','VideString','<<Tous les établissements>>');
    setControlProperty('_VALNIV','DataType','TTETABLISSEMENT');
    if Sender <> nil then SetControlText('_VALNIV','');
    setControlProperty('_VALNIV','Enabled',True);
    setControlProperty('LBL_VALNIV','Enabled',True);
  end else if st = 'CON' then
  begin
    setControlProperty('_VALNIV','Vide',False);
    setControlProperty('_VALNIV','VideString','');
    setControlProperty('_VALNIV','DataType','PGCONVENTIONS');
    if Sender <> nil then SetControlText('_VALNIV','');
    setControlProperty('_VALNIV','Enabled',True);
    setControlProperty('LBL_VALNIV','Enabled',True);
  end else begin
    // L'affectation de '' ne fonctionne pas, la dernière tablette associé reste
    setControlProperty('_VALNIV','DataType','');
    if Sender <> nil then setControlText('_VALNIV','...');
    setControlProperty('_VALNIV','Enabled',False);
    setControlProperty('LBL_VALNIV','Enabled',False);
  end;
  if Sender <> nil then BtnValider.Enabled := True; //Autorisation;
end;


procedure TOF_PGTABLESDYNADUP.OnChangeDTVALID(Sender: TObject);
begin
  // Mettre un message si il existe déjà une table dynamique dont la date de validité est postérieur
  BtnValider.Enabled := True; //Autorisation;
end;

procedure TOF_PGTABLESDYNADUP.OnChangeVALNIV(Sender: TObject);
begin
  BtnValider.Enabled := True; //Autorisation;
  if GetControlText('_VALNIV')<> '' then
    VALNIV  := GetControlText('_VALNIV')
  else
    VALNIV  := '...';
end;

Initialization
  registerclasses ( [ TOF_PGTABLESDYNADUP ] ) ; 
end.
