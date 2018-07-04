{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 21/03/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : AFVALINDICE (AFVALINDICE)
Mots clefs ... : TOM;AFVALINDICE
*****************************************************************}
Unit uTomAfValIndice ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
     Fiche,
{$Else}
     MainEagl,         
     eFiche,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     UTob,HTB97, HDB ;

Type
  TOM_AFVALINDICE = Class (TOM)

      procedure OnNewRecord                ; override ;
      procedure OnDeleteRecord             ; override ;
      procedure OnUpdateRecord             ; override ;
      procedure OnAfterUpdateRecord        ; override ;
      procedure OnLoadRecord               ; override ;
      procedure OnChangeField ( F: TField) ; override ;
      procedure OnArgument ( S: String )   ; override ;
      procedure OnClose                    ; override ;

      procedure OnCancelRecord             ; override ;
      procedure CbIndiceSuivChange(sender : Tobject) ;
      procedure CbIndiceChange(sender : Tobject) ;
      procedure CbPubChange(sender : Tobject) ;
      procedure CbPubSuivChange(sender : Tobject) ;
      procedure bCalculeCoefClick(sender : Tobject) ;
      procedure BEFFACEINDICESUIVANTClick(sender : Tobject) ;
      procedure BEFFACEPUBLICATIONClick(sender : Tobject) ;
      procedure BDUPLICATIONClick(sender : Tobject) ;
      procedure AFV_INDDATEVALExit(sender : Tobject) ;

    private
      bCalculeCoef          : TToolbarButton97 ;
      BEFFACEINDICESUIVANT  : TToolbarButton97 ;
      BEFFACEPUBLICATION    : TToolbarButton97 ;
      BDUPLICATION          : TToolbarButton97 ;
      CbIndiceSuiv          : THDBValComboBox ;
      LbIndiceSuiv          : THLabel ;
      CbIndice              : THDBValComboBox ;
      LbIndice              : THLabel ;
      CbPub                 : THDBValComboBox ;
      LbPub                 : THLabel ;
      CbPubSuiv             : THDBValComboBox ;
      LbPubSuiv             : THLabel ;
    end;


 const
  	TexteMsg: array[1..5] of string 	= (
          {1}        'le code de la publication est obligatoire.',
          {2}        'le numéro de la publication est obligatoire.',
          {3}        'la date de début d''application est obligatoire.',
          {4}        'La valeur de l''indice est obligatoire.',
          {5}        'la date de publication est obligatoire.'
          );

procedure AglLanceFicheAFVALINDICE(cle,Action : string ) ;
Implementation

procedure TOM_AFVALINDICE.OnNewRecord ;
begin
  Inherited ;
  SetField('AFV_INDDATEFIN', iDate2099);
  SetControlText('AIN_INDDATECREATION', '');
  SetControlText('AIN_INDDATECREATIONSUIV', '');
  SetField('AFV_DEFINITIF', 'X');
  SetControlText('AFV_INDMAJ_LIB', 'Saisie');
  SetField('AFV_INDMAJ', 'SAI');
  SetField('AFV_DATEMAJ', date);
end ;

procedure TOM_AFVALINDICE.OnDeleteRecord ;
begin
  Inherited;
  SetControlText('LBLIBELLEINDICE', '');
  SetControlText('LBLIBELLEPUBLICATION', '');
  SetControlText('LBLIBELLEINDICESUIVANT', '');
  SetControlText('LBLIBELLEPUBLICATIONSUIV', '');
  SetControlText('AIN_INDDATECREATION', '');
  SetControlText('AFV_INDMAJ_LIB', '');
  SetControlText('AIN_INDDATECREATIONSUIV', '');
end;

procedure TOM_AFVALINDICE.OnUpdateRecord ;
Var
  erreur  : integer;
  st      : string;

begin
  erreur:=0 ;
  if getcontroltext('AFV_PUBCODE')='' then  erreur := 1
   else if getcontroltext('AFV_INDNUMPUB')='' then erreur := 2
     else if (getcontroltext('AFV_INDDATEVAL')='') or
              (strToDate(getcontroltext('AFV_INDDATEVAL')) = iDate1900) then erreur := 3
       else if (getcontroltext('AFV_INDVALEUR')='') or
              (getcontroltext('AFV_INDVALEUR')='0') then erreur := 4
         else if (getcontroltext('AFV_INDDATEPUB')='') or
              (strToDate(getcontroltext('AFV_INDDATEPUB')) = iDate1900) then erreur := 5 ;
   
  if erreur>0 then
  begin
    LastError:=erreur ;
    LastErrorMsg := TexteMsg[LastError];
  end 
  Else
  begin
    st:='update AFINDICE set AIN_INDDATEFIN="'+usdatetime(strtodate(getcontroltext('AFV_INDDATEFIN')))+'"  where AIN_INDCODE="'+getcontroltext('AFV_INDCODE')+'"';
    ExecuteSQL(St);
  end;
  Inherited;
end;

procedure TOM_AFVALINDICE.OnAfterUpdateRecord ;
begin
  Inherited;
end;
 
procedure TOM_AFVALINDICE.OnLoadRecord ;
begin
  Inherited;
  CbIndiceSuivChange(nil) ;
  CbIndiceChange(nil) ;
  CbPubChange(nil) ;
  CbPubSuivChange(nil) ;
  SetFocusControl('AFV_INDCODE');
  SetControlText('AFV_INDMAJ_LIB', RechDom('AFTMAJINDICE', GetField('AFV_INDMAJ'), false));
end;

procedure TOM_AFVALINDICE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_AFVALINDICE.CbIndiceSuivChange(sender : Tobject) ;
var
  vSt : String;
  vQr : TQuery;

begin
  inherited;
  if CbIndiceSuiv.value <> '' then
    LbIndiceSuiv.Caption:=RECHDOM('AFTINDICE_LIB',CbIndiceSuiv.value,false)
  else
    LbIndiceSuiv.Caption:='' ;

  vSt := 'SELECT AIN_INDDATECREA FROM AFINDICE WHERE AIN_INDCODE = "' + CbIndiceSuiv.value + '"';
  vQr := nil;
  Try
    vQR := OpenSql(vSt, True);
     if Not vQR.Eof then
     begin
       setControlText('AIN_INDDATECREATIONSUIV', vQr.FindField('AIN_INDDATECREA').asString);
     end;
  Finally
    ferme(vQr);
  End;
end;

procedure TOM_AFVALINDICE.CbIndiceChange(sender : Tobject) ;
var
  vSt : String;
  vQr : TQuery;

begin
  if CbIndice.value<> '' then
    LbIndice.Caption:=RECHDOM('AFTINDICE_LIB',CbIndice.value,false)
  else
    LbIndice.Caption:='' ;

  vSt := 'SELECT AIN_INDDATECREA FROM AFINDICE WHERE AIN_INDCODE = "' + GetControlText('AFV_INDCODE') + '"';
  vQr := nil;
  Try
    vQR := OpenSql(vSt, True);
    if Not vQR.Eof then
      setControlText('AIN_INDDATECREATION', vQr.FindField('AIN_INDDATECREA').asString);
  Finally
    ferme(vQr);
  End;

  // on prend la derniere date d'application saisie
  // attention ceci est egalement utilisé au load
  // on supprime le chargement de pub_code si déjà correctement chargé pour
  // ne pas passer en mode modification
  vSt := 'SELECT AFV_INDVALEUR, AFV_INDNUMPUB, AFV_PUBCODE FROM AFVALINDICE ';
  vSt := vSt + ' WHERE AFV_INDCODE = "' + GetControlText('AFV_INDCODE') + '"';
  vSt := vSt + ' ORDER BY AFV_INDDATEVAL DESC';

  vQr := nil;
  Try
    vQR := OpenSql(vSt, True);
    if Not vQR.Eof then
    begin
      if GetControlText('AFV_INDNUMPUB') = '' then
      begin
        setControlText('AFV_INDVALEUR', vQr.FindField('AFV_INDVALEUR').asString);
        setControlText('AFV_INDNUMPUB', vQr.FindField('AFV_INDNUMPUB').asString);
      end;            
      if vQr.FindField('AFV_PUBCODE').asString <> GetControlText('AFV_PUBCODE') then
        setControlText('AFV_PUBCODE', vQr.FindField('AFV_PUBCODE').asString);
    end;                                
  Finally
    ferme(vQr);
  End;
end;

procedure TOM_AFVALINDICE.CbPubChange(sender : Tobject) ;
begin
  if CbPub.value<> '' then
    LbPub.Caption:=RECHDOM('AFPUBLICATION_LIB',CbPub.value,false)
  else
    LbPub.Caption:='' ;
end ;

procedure TOM_AFVALINDICE.CbPubSuivChange(sender : Tobject) ;
begin
  if CbPubSuiv.value<> '' then
    LbPubSuiv.Caption:=RECHDOM('AFPUBLICATION_LIB',CbPubSuiv.value,false)
  else
    LbPubSuiv.Caption:='' ;
end ;


procedure TOM_AFVALINDICE.BEFFACEPUBLICATIONClick(sender : Tobject) ;
begin
  setcontroltext('AFV_PUBCODESUIV','') ;
end ;

procedure TOM_AFVALINDICE.BDUPLICATIONClick(sender : Tobject);
var
  fTobDuplic : Tob;
begin

  fTobDuplic := Tob.create('AFVALINDICE',nil,-1);
  Try
    fTobDuplic.PutValue('AFV_PUBCODE', GetField('AFV_PUBCODE'));
    fTobDuplic.PutValue('AFV_INDCODE', GetField('AFV_INDCODE'));
    fTobDuplic.PutValue('AFV_INDCOMMENT', GetField('AFV_INDCOMMENT'));
    fTobDuplic.PutValue('AFV_INDDATEFIN', GetField('AFV_INDDATEFIN'));
    fTobDuplic.PutValue('AFV_INDCODESUIV', GetField('AFV_INDCODESUIV'));
    fTobDuplic.PutValue('AFV_COEFPASSAGE', GetField('AFV_COEFPASSAGE'));
    fTobDuplic.PutValue('AFV_PUBCODESUIV', GetField('AFV_PUBCODESUIV'));
    fTobDuplic.PutValue('AFV_COEFRACCORD', GetField('AFV_COEFRACCORD'));

    TFFiche(Ecran).Binsert.Click;

    SetControlText('AFV_PUBCODE', fTobDuplic.GetValue('AFV_PUBCODE'));
    SetControlText('AFV_INDCODE', fTobDuplic.GetValue('AFV_INDCODE'));
    SetControlText('AFV_INDCOMMENT', fTobDuplic.GetValue('AFV_INDCOMMENT'));
    SetControlText('AFV_INDDATEFIN', fTobDuplic.GetValue('AFV_INDDATEFIN'));
    SetControlText('AFV_INDCODESUIV', fTobDuplic.GetValue('AFV_INDCODESUIV'));
    SetControlText('AFV_COEFPASSAGE', fTobDuplic.GetValue('AFV_COEFPASSAGE'));
    SetControlText('AFV_PUBCODESUIV', fTobDuplic.GetValue('AFV_PUBCODESUIV'));
    SetControlText('AFV_COEFRACCORD', fTobDuplic.GetValue('AFV_COEFRACCORD'));

  Finally
    fTobDuplic.Free;
  end;
end;

procedure TOM_AFVALINDICE.BEFFACEINDICESUIVANTClick(sender : Tobject) ;
begin
  setcontroltext('AFV_INDCODESUIV','') ;
  setcontroltext('AIN_INDDATECREATIONSUIV','') ;
end ;

procedure TOM_AFVALINDICE.bCalculeCoefClick(sender : Tobject) ;
Var StRequete : String ;
   MaTOB:TOB;
   MaRequete : TQuery ;
   ValIndiceSuivant : Double ;

begin
  ValIndiceSuivant:=0 ;
//  NextPrevControl(self.Ecran) ;
  StRequete:='SELECT AFV_INDVALEUR,MAX(AFV_INDDATEVAL) FROM AFVALINDICE  WHERE ' ;
  StRequete:=StRequete+' ((AFV_INDDATEVAL <="'+USDATETIME(strtodate(GetField('AFV_INDDATEFIN')))+'")'  ;
  StRequete:=StRequete+' AND (AFV_INDCODE ="'+GetControlText('AFV_INDCODESUIV')+'")) Group By AFV_INDVALEUR'  ;
  MaRequete:=Nil ;
  MaTOB:=TOB.Create('',nil,-1);
  try
    MaRequete := OpenSQL(StRequete, TRUE);
    MaTOB.LoadDetailDB('','','',MaRequete,false) ;
    if MaTOB.Detail.count>0 then
    begin
      ValIndiceSuivant:=MaTOB.Detail[0].GetValue('AFV_INDVALEUR') ;
    end ;
  finally
    Ferme(MaRequete);
    MaTOB.Free;
  end ;
  if ValIndiceSuivant<>0 then setcontroltext('AFV_COEFPASSAGE',floattostr(strtofloat(getcontroltext('AFV_INDVALEUR'))/ValIndiceSuivant)) ;
  SetFocusControl('AFV_COEFRACCORD') ;
//  NextPrevControl(self.Ecran) ;
end ;

procedure TOM_AFVALINDICE.OnArgument ( S: String ) ;
begin
  Inherited ;
  CbIndiceSuiv:=THDBValComboBox(GetControl('AFV_INDCODESUIV')) ;
  LbIndiceSuiv:=THLabel(GetControl('LBLIBELLEINDICESUIVANT')) ;
  CbIndiceSuiv.OnChange:=CbIndiceSuivChange ;

  CbIndice:=THDBValComboBox(GetControl('AFV_INDCODE')) ;
  LbIndice:=THLabel(GetControl('LBLIBELLEINDICE')) ;
  CbIndice.OnChange:=CbIndiceChange ;

  CbPub:=THDBValComboBox(GetControl('AFV_PUBCODE')) ;
  LbPub:=THLabel(GetControl('LBLIBELLEPUBLICATION')) ;
  CbPub.OnChange:=CbPubChange ;

  CbPubSuiv:=THDBValComboBox(GetControl('AFV_PUBCODESUIV')) ;
  LbPubSuiv:=THLabel(GetControl('LBLIBELLEPUBLICATIONSUIV')) ;
  CbPubSuiv.OnChange:=CbPubSuivChange ;

  bCalculeCoef:=TToolbarButton97(GetControl('BCALCULECOEF')) ;
  bCalculeCoef.OnClick:=bCalculeCoefClick ;

  BEFFACEINDICESUIVANT:=TToolbarButton97(GetControl('BEFFACEINDICESUIVANT')) ;
  BEFFACEINDICESUIVANT.OnClick:=BEFFACEINDICESUIVANTClick ;

  BEFFACEPUBLICATION:=TToolbarButton97(GetControl('BEFFACEPUBLICATION')) ;
  BEFFACEPUBLICATION.OnClick:=BEFFACEPUBLICATIONClick ;

  BDUPLICATION:=TToolbarButton97(GetControl('BDUPLICATION')) ;
  BDUPLICATION.OnClick:=BDUPLICATIONClick ;

  THDBEdit(GetControl('AFV_INDDATEVAL')).OnExit := AFV_INDDATEVALExit;
end ;


procedure TOM_AFVALINDICE.AFV_INDDATEVALExit(sender : Tobject) ;
begin
  Inherited ;
  SetField('AFV_INDDATEPUB', getField('AFV_INDDATEVAL'));
end;

procedure TOM_AFVALINDICE.OnClose ;
begin
  Inherited ;
end ;
 
procedure TOM_AFVALINDICE.OnCancelRecord ;
begin
  Inherited ;
end ;                             

procedure AglLanceFicheAFVALINDICE(cle,Action : string ) ;
begin
  AglLanceFiche('AFF','AFVALINDICE','',cle,Action);
end ;

Initialization
  registerclasses ( [ TOM_AFVALINDICE ] ) ;
end.
