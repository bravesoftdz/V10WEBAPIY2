{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 18/03/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : RHCOMPETRESSOURCE (RHCOMPETRESSOURCE)
Mots clefs ... : TOM;RHCOMPETRESSOURCE
*****************************************************************
PT1 24/11/2004 V_60 JL Contrôle code compétence <> ''
}
Unit RHCOMPETRESSOURCE_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$else}
     eFiche,
     eFichList,
{$ENDIF}
     sysutils,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     {$IFNDEF PAIEGRH}
     wTom,
     {$ELSE}
     EntPaie,
     {$ENDIF}
     uTableFiltre,
     SaisieList,
     Utob;


Type
  {$IFNDEF PAIEGRH}
  TOM_RHCOMPETRESSOURCE = Class (twTOM)
  {$ELSE}
  TOM_RHCOMPETRESSOURCE = Class (TOM)
  {$ENDIF}
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    {$IFNDEF PAIEGRH}
    Protected
       function RecordIsValid : Boolean; override;
    {$ENDIF}
    Private
       lGRHSerialise : Boolean;
       NatureSaisie : String;
       TF : TTableFiltre;
    end ;

Implementation
{$IFNDEF PAIEGRH}
uses wCommuns;
{$ENDIF}

Const
  TexteMessage: array[1..6] of string = (
	    {1}  'La date de fin de validité doit être postérieur à la date de validation.' + #13 + ' Modifier l''une des deux date.',
	    {2}  'La date de fin de suspension doit être postérieur à la date de suspension.' + #13 + ' Modifier l''une des deux date.',
      {3}  'Le motif de suspension n''est pas renseigné.',
      {4}  'Une compétence ne peut être validée par la ressource à qui elle est attribuée.',
      {5}  'Le code compétence n''est pas renseigné.',
      {6}  'Cette compétence a déjà été affectée à cette ressource.' + #13 + ' Saisir une autre compétence'
		  );

      ErrFourchetteDate     = 1;
      ErrFourchetteDateSusp = 2;
      ErrMotifSusp          = 3;
      ErrValideParLuiMeme   = 4;
      ErrCompetenceOblig    = 5;
      ErrCompetexiste       = 6;

function EstCompetent(Ressource, Competence : String) : Boolean;
var sRequ : String;
begin
  srequ := 'Select PCH_RESSOURCE FROM RHCOMPETRESSOURCE '
            + 'WHERE PCH_RESSOURCE="'+Ressource+'" AND PCH_COMPETENCE="'+Competence+'"';
  result := ExisteSql(srequ);
end;

{$IFNDEF PAIEGRH}
function TOM_RHCOMPETRESSOURCE.recordIsValid : Boolean;
begin
  Result := false;

  { Contrôle bas niveau }
  if not inherited RecordIsvalid then exit;

  { Contrôles écran }
  if Assigned(ecran) then
  begin
    // Pas d'écran
  end;

  LastError := 0;
  if GetDateTime('PCH_DATEVALIDATION') > GetDateTime('PCH_DATEFINVALID') then
    LastError := ErrFourchetteDate;
  if GetDateTime('PCH_DEBUTSUSP') > GetDateTime('PCH_FINSUSP') then
    LastError := ErrFourchetteDateSusp;
  if GetBoolean('PCH_SUSPENSION') AND (Trim(GetString('PCH_MOTIFSUSPCOMPT'))='') then
    LastError := ErrMotifSusp;
  if GetString('PCH_VALIDEPAR') = GetString('PCH_RESSOURCE') then
    LastError := ErrValideParLuiMeme;
  if (Trim(GetString('PCH_COMPETENCE'))='')  then
    LastError := ErrCompetenceOblig;
  If ((IKC = 'C') OR (IKC = 'D')) and EstCompetent(GetString('PCH_RESSOURCE'),GetString('PCH_COMPETENCE')) then
    LastError := ErrCompetexiste;

  if LastError > 0 then
  begin
    if LastError <> ErrCompetenceOblig then
      LastErrorMsg := TexteMessage[LastError];

    if Assigned(ecran) then
    begin
      // C.B 05/05/2006
      // les messages d'erreurs étaient affichés 2 fois  
      //PGIBox(LastErrorMsg);
      case LastError of
        ErrFourchetteDate     : SetFocusControl('PCH_DATEFINVALID');
        ErrFourchetteDateSusp : SetFocusControl('PCH_FINSUSP');
        ErrMotifSusp          : SetFocusControl('PCH_MOTIFSUSPCOMPT');
        ErrValideParLuiMeme   : SetFocusControl('PCH_VALIDEPAR');
        ErrCompetenceOblig    : SetFocusControl('PCH_COMPETENCE');
        ErrCompetExiste       : SetFocusControl('PCH_COMPETENCE');
      end;
    end
    else
    begin
      fTob.AddChampSupValeur('Error', LastErrorMsg, false);
      case LastError of
        ErrFourchetteDate     : fTob.AddChampSupValeur('FieldError', 'PCH_DATEFINVALID'  , false);
        ErrFourchetteDateSusp : fTob.AddChampSupValeur('FieldError', 'PCH_FINSUSP'       , false);
        ErrMotifSusp          : fTob.AddChampSupValeur('FieldError', 'PCH_MOTIFSUSPCOMPT', False);
        ErrValideParLuiMeme   : fTob.AddChampSupValeur('FieldError', 'PCH_VALIDEPAR'     , False);
        ErrCompetenceOblig    : fTob.AddChampSupValeur('FieldError', 'PCH_COMPETENCE'    , False);
        ErrCompetExiste       : fTob.AddChampSupValeur('FieldError', 'PCH_COMPETENCE'    , False);
      end;
      if V_Pgi.Sav then PgiInfo(LastErrorMsg, 'RHCOMPETRESSOURCE_Tom.RecordIsValid');
    end;
  end;
  Result := LastError=0;
end ;
{$ENDIF}


procedure TOM_RHCOMPETRESSOURCE.OnNewRecord ;
{$IFDEF PAIEGRH}
var
  TypeInt,Salarie : String;
  Q : TQuery;
{$ENDIF PAIEGRH}
begin
  Inherited ;
{$IFDEF PAIEGRH}
  If GetControl('VALIDATION') <> Nil then
  begin
    If GetControlText('VALIDATION') <> '' then SetField('PCH_VALIDEPAR',GetControlText('VALIDATION'));
  end;
  TypeInt := TF.TOBFiltre.GetValue('PSI_TYPEINTERIM');
  Salarie := TF.TOBFiltre.GetValue('PSI_INTERIMAIRE');
  SetField('PCH_TYPERESSOURCE',TypeInt);
  SetField('PCH_SALARIE',Salarie);
  SetField('PCH_DATEVALIDATION',Date);
  SetField('PCH_DATEFINVALID',IDate1900);
  SetField('PCH_DEBUTSUSP',IDate1900);
  SetField('PCH_FINSUSP',IDate1900);
  Q := OpenSQL('SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_SALARIE="'+Salarie+'" AND ARS_TYPERESSOURCE="'+TypeInt+'"',True);
  If Not Q.Eof then SetField('PCH_RESSOURCE',Q.FindField('ARS_RESSOURCE').AsString);
  Ferme(Q);
{$ELSE}
  // 22/02/2005 initialisation de la fiche liste WCOMPETRESRCE
  if (Ecran<>nil) and (Ecran is TFSaisieList) then
  begin
    SetField('PCH_DATEVALIDATION', Date);
    SetField('PCH_DATEFINVALID', iDate2099);
    SetField('PCH_COMPETENCE','');//AB-200509-FQ12450
  end;
{$ENDIF}
end ;

procedure TOM_RHCOMPETRESSOURCE.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_RHCOMPETRESSOURCE.OnUpdateRecord ;
var
  {$IFNDEF PAIEGRH}
    ARSDATAS : MyArrayValue;
  {$ELSE !PAIEGRH}
    DD,DF : TDateTime;
  {$ENDIF !PAIEGRH}
  IMax : Integer;
  Q : TQuery;
begin
  Inherited;                 
  {$IFNDEF PAIEGRH}
  ARSDATAS := nil;
  if RecordIsValid then
  begin
    if Ikc = 'C' then
  	begin
      { Jeton }
      //SetInteger('PCH_RANG', WSetJeton('PCH'));
      ARSDATAS := wGetSQLFieldsValues(['ARS_SALARIE','ARS_TYPERESSOURCE'], 'RESSOURCE', 'ARS_RESSOURCE="'+getString('PCH_RESSOURCE')+'"');
      SetString('PCH_TYPERESSOURCE', String(ARSDATAS[1]));
      SetString('PCH_SALARIE', String(ARSDATAS[0]));
    end;
  end
  else Exit;  //AB-20050519
  {$ELSE}
  LastError := 0;
  //DEBUT PT1
  If GetField('PCH_COMPETENCE') = '' then
  begin
        LastError := 1;
       PGIBox(TexteMessage[5],Ecran.Caption);
       SetFocusControl('PCH_COMPETENCE');
       Exit;
  end;
  //FIN PT1
  DD := GetField('PCH_DATEVALIDATION');
  DF := GetField('PCH_DATEFINVALID');
  If (DD > DF) and (DF> IDate1900) then
  begin
       LastError := 1;
       PGIBox(TexteMessage[1],Ecran.Caption);
       SetFocusControl('PCH_DATEVALIDATION');
       Exit;
  end;
  DD := GetField('PCH_DEBUTSUSP');
  DF := GetField('PCH_FINSUSP');
  If (DD > DF) and (DF> IDate1900) then
  begin
       LastError := 1;
       PGIBox(TexteMessage[2],Ecran.Caption);
       SetFocusControl('PCH_DEBUTSUSP');
       Exit;
  end;
  {$ENDIF}
  if DS.State in [dsInsert] then
  begin
        Q:=OpenSQL('SELECT MAX(PCH_RANG) FROM RHCOMPETRESSOURCE WHERE PCH_RESSOURCE="'+GetField('PCH_RESSOURCE')+'"',TRUE) ;
        if Not Q.EOF then imax:=Q.Fields[0].AsInteger+1 else iMax:=1 ;
        Ferme(Q) ;
        SetField ('PCH_RANG',Imax);
  end;
end ;

procedure TOM_RHCOMPETRESSOURCE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_RHCOMPETRESSOURCE.OnLoadRecord ;
begin
{$IFNDEF PAIEGRH}
  SetControlEnabled('PCH_MOTIFSUSPCOMPT',tCheckBox(GetControl('PCH_SUSPENSION')).checked and not lGRHSerialise);
  SetControlEnabled('PCH_DEBUTSUSP',tCheckBox(GetControl('PCH_SUSPENSION')).checked and not lGRHSerialise);
  SetControlEnabled('PCH_FINSUSP',tCheckBox(GetControl('PCH_SUSPENSION')).checked and not lGRHSerialise);
{$ENDIF}
end ;

procedure TOM_RHCOMPETRESSOURCE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
  {$IFDEF PAIEGRH}
  if F.FieldName = 'PCH_SUSPENSION' then
  begin
    SetControlEnabled('PCH_MOTIFSUSPCOMPT',tCheckBox(GetControl('PCH_SUSPENSION')).checked);
    SetControlEnabled('PCH_DEBUTSUSP',tCheckBox(GetControl('PCH_SUSPENSION')).checked);
    SetControlEnabled('PCH_FINSUSP',tCheckBox(GetControl('PCH_SUSPENSION')).checked);
  end;
  {$ENDIF}
end ;

procedure TOM_RHCOMPETRESSOURCE.OnArgument ( S: String ) ;
begin
  Inherited ;
  lGRHSerialise := True; // A initialiser quand le module GRH sera sérialisé.
  {$IFDEF PAIEGRH}
  lGRHSerialise := VH_PAIE.PgSeriaCompetence;
  {$ENDIF}
  TF  :=  TFSaisieList(Ecran).LeFiltre;
  NatureSaisie := ReadTokenPipe(S,';');
end ;

procedure TOM_RHCOMPETRESSOURCE.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_RHCOMPETRESSOURCE.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_RHCOMPETRESSOURCE ] ) ;
end.



