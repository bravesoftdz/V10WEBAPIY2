unit UtofRessource_Mul;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,HCtrls,UTOF,
     DicoBTP,
      // grc
      HQry,M3FP,HStatus,
{$IFDEF EAGLCLIENT}
     eMul,Maineagl,
{$ELSE}
     Mul,HDB,FE_Main, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db,
{$ENDIF}
     Ent1,EntGC,
{$ifdef AFFAIRE}
     TraducAffaire,
{$endif}
     UtilPGI,
     HEnt1,
     UtilGc,
     Menus,
     ConfidentAffaire,
     UTOFAFTRADUCCHAMPLIBRE,
     HTB97,
     utob;

Type
     TOF_RESSOURCE_MUL = Class (TOF_AFTRADUCCHAMPLIBRE)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnUpdate ; Override;
        procedure ColTauxInvisibles;
        // grc
        Function RameneListeRessources : String;
     private
        StFiltre : string ;
        BtInsert : TToolbarbutton97;
        TypeMultiRess	: THMultiValComboBox;
        Fonction      : THValComboBox;
        TypeRessource : THValComboBox;
        Fliste 	 : THDbGrid;
        Action				: String;
        procedure AfterShow;
    {$IFDEF BTP}
        procedure ConstituePlusTypeRessource(TheTypes: string);
    {$ENDIF}
		    Procedure BInsert_OnClick(Sender: TObject);
        Procedure CreateSalarie(Sender: TObject);
        Procedure CreateSsTraitance(Sender: TObject);
				Procedure CreateIterim(Sender: TObject);
				Procedure CreateMateriel(Sender: TObject);
				Procedure CreateOutil(Sender: TObject);
				Procedure CreateLocation(Sender: TObject);
				Procedure CreateAutre(Sender: TObject);
        //
        procedure FlisteDblClick(Sender: TObject);
    		procedure CreationRessource(TypeRessource: String);
     END ;

// grc
Function AGLRameneListeRessources(parms:array of variant; nb: integer ) : variant ;
Procedure AFLanceFiche_Mul_Ressource;
Procedure AFLanceFiche_Mul_RessourceLine(Argument : String);
Function AFLanceFiche_Rech_Ressource(Range,Argument:string):variant;
//


implementation
uses
  Paramsoc
  , uTOFComm
  , FormsName
  ;

procedure TOF_RESSOURCE_MUL.OnArgument(stArgument : String );
// Ajout PL le 28/02/02
var      Critere, Champ, valeur : string;
         x : integer;
// Fin Ajout PL le 28/02/02
         F : TForm;
begin
fMulDeTraitement  := true;
Inherited;
fTableName := 'RESSOURCE';

UpdateCaption(Ecran);
if VH_GC.GCIfDefCEGID then
begin
  SetControlVisible('PComplement',False); SetControlVisible('PCompl',False);
  SetControlVisible('BInsert',False);
end;

	if not AffichageValorisation then
  begin
     SetControlVisible('ARS_TAUXREVIENTUN', false);
     SetControlVisible('ARS_TAUXREVIENTUN_', false);
     SetControlVisible('TARS_TAUXREVIENTUN', false);
     SetControlVisible('TARS_TAUXREVIENTUN_', false);
  end;

  if (THDbGrid (GetControl('FLISTE')) <> nil) (*and (ecran.name<>'BTRESSRECH_MUL')*) then
  Begin
	   Fliste := THDbGrid (GetControl('FLISTE'));
     Fliste.OnDblClick := FlisteDblClick;
  end;

 	BtInsert := TToolbarButton97 (ecran.findcomponent('Binsert'));

  if Tpopupmenu(GetControl('POPCREATERESS')) <> nil then
  Begin
     TMenuItem(GetControl ('CREASAL')).onclick := CreateSalarie;
     TMenuItem(GetControl ('CREASSTRT')).onclick := CreateSsTraitance;
     TMenuItem(GetControl ('CREAINTERIM')).onclick := CreateIterim;
     TMenuItem(GetControl ('CREAMAT')).onclick := CreateMateriel;
     TMenuItem(GetControl ('CREAOUTIL')).onclick := CreateOutil;
     TMenuItem(GetControl ('CREALOC')).onclick := CreateLocation;
     if assigned(GetControl ('CREAAUT')) then TMenuItem(GetControl ('CREAAUT')).onclick := CreateAutre;
  End;

  if Ecran.Name = 'BTRESSRECH_MUL' then
  Begin
	   TypeMultiRess := THMultiValComBoBox(Ecran.FindComponent('ARS_TYPERESSOURCE'));
	   if Tpopupmenu(GetControl('POPCREATERESS')) <> nil then
     Begin
        BtInsert.DropdownMenu := Tpopupmenu(GetControl('POPCREATERESS'));
        BtInsert.DropDownAlways := true;
        BtInsert.Width := 35;
     end;
  end else
  Begin
  	 TypeRessource := THValComBoBox(Ecran.FindComponent('ARS_TYPERESSOURCE'));

//Non en line
	   if Tpopupmenu(GetControl('POPCREATERESS')) <> nil then
     Begin
    		BtInsert.DropDownAlways := true;
		    BtInsert.DropdownMenu := Tpopupmenu(GetControl('POPCREATERESS'));
		    BtInsert.Width := 35;
     end;
//	   BTinsert.OnClick := BInsert_OnClick;
	end;

  Fonction := THValComBoBox(GetControl('ARS_FONCTION1'));

  // Ajout PL le 28/02/02 pour afficher la partie du code ressource saisie avant de lancer le mul
  Critere:=(Trim(ReadTokenSt(stArgument)));
  While (Critere <>'') do
    BEGIN
    if Critere<>'' then
        BEGIN
        X:=pos(':',Critere);
        if (x=0 ) or ( copy(Critere,1,6) = 'FILTRE') then  X:=pos('=',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;
        if Champ = 'ARS_RESSOURCE'    then
            begin
            SetControlText('ARS_RESSOURCE',Valeur);
            end ;
        if Champ = 'FILTRE'    then
            StFiltre:=Valeur;
        if Champ = 'ACTION' Then
           Action := Valeur;
{$IFDEF BTP}
        // constitue le + du type de ressource pour n'afficher que les type de ressource désiré
        // passage d'un certain nombre de valeur séparateur "," et exclusion '-'
        if trim(Champ) = 'ARS_FONCTION1' then
           begin
          if Fonction <> nil then Fonction.Value := Trim(Valeur);
        end;
        if trim(Champ) = 'TYPERESSOURCE' then
        begin
           ConstituePlusTypeRessource(valeur);
           end;
{$ENDIF}
        END;
     Critere:=(Trim(ReadTokenSt(stArgument)));
     END;

  BtInsert := TToolbarButton97 (ecran.findcomponent('Binsert'));
  BtInsert.OnClick := Binsert_OnClick;
// Fin Ajout PL le 28/02/02
  F := TForm (Ecran);
  if F.Name='RTRESSOURCE_TL' then
  	 begin
     Tfmul(Ecran).FiltreDisabled:=true;
     TFmul(Ecran).OnAfterFormShow := AfterShow;
     end;

//Gestion des ongleets en LINE
//uniquement en line
//	SetControlProperty('PComplement', 'TabVisible', False);
//  SetControlProperty('PCOMP', 'TabVisible', False);

//Gestion des ongleets en LINE
//uniquement en line
//	SetControlProperty('PComplement', 'TabVisible', False);
//  SetControlProperty('PCOMP', 'TabVisible', False);

End;

procedure TOF_RESSOURCE_MUL.ColTauxInvisibles;
{$IFNDEF EAGLCLIENT}
var
i:integer;
{$ENDIF}
Begin

{$IFDEF EAGLCLIENT}
//AFAIREEAGL
{$ELSE}
For i:=0 to TFMul(Ecran).FListe.Columns.Count-1 do
    Begin
    If (TFMul(Ecran).FListe.Columns[i].FieldName = 'ARS_TAUXREVIENTUN')
        OR (TFMul(Ecran).FListe.Columns[i].FieldName = 'ARS_PVHT')
        OR (TFMul(Ecran).FListe.Columns[i].FieldName = 'ARS_PVTTC')
        OR (TFMul(Ecran).FListe.Columns[i].FieldName = 'ARS_TAUXUNIT')
        Then
        TFMul(Ecran).FListe.columns[i].visible:=False ;
      end;
{$ENDIF}
End;

procedure TOF_RESSOURCE_MUL.OnUpdate;
Begin
if not AffichageValorisation then
    begin
    ColTauxInvisibles;
    end ;

{$ifdef AFFAIRE}
  {$IFDEF EAGLCLIENT}
TraduitAFLibGridSt(TFMul(Ecran).FListe);
  {$ELSE}
TraduitAFLibGridDB(TFMul(Ecran).FListe);
  {$ENDIF}
{$ENDIF}
End;

// grc
Function AGLRameneListeRessources( parms: array of variant; nb: integer ) : variant;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_RESSOURCE_MUL) then result:=TOF_RESSOURCE_MUL(TOTOF).RameneListeRessources else exit;
end;

Function TOF_RESSOURCE_MUL.RameneListeRessources : String;
var F : TFMul;
    Q : THQuery;
    i : integer;
{$IFDEF EAGLCLIENT}
       L : THGrid;
{$ELSE}
       L : THDBGrid;
{$ENDIF}
    code : string;
begin
Result := '';
F:=TFMul(Ecran);
L:= F.FListe;
Q:= F.Q;

{$IFDEF EAGLCLIENT}
if F.bSelectAll.Down then
  if not F.Fetchlestous then
    begin
    F.bSelectAllClick(Nil);
    F.bSelectAll.Down := False;
    exit;
    end;
{$ENDIF}


if L.AllSelected then
    begin
    Q.First ;
    While not Q.Eof do
      begin
      code:=Q.FindField('ARS_RESSOURCE').asstring ;
      if Result = '' then Result:=code else Result:=Result+';'+code;
      // Pas de solution pour l'instant pour la next. le select all ne fonctionne pas en eagl (ne ramène pas tout)
      // et en agl on peut récupérer toutes les ressources
      // ceci dit, il ne peut pas y avoir 1 million de ressources....
      Q.Next ;
      end;
      L.AllSelected:=False;
    end
else
  if F.FListe.NbSelected <> 0 then
    begin
    InitMove(L.NbSelected,'');
    for i:=0 to L.NbSelected-1 do
      begin
      MoveCur(False);
      L.GotoLeBookmark(i);
      {$IFDEF EAGLCLIENT}
      Q.TQ.Seek(L.Row-1) ;
			{$ENDIF}
      code:=TFmul(Ecran).Q.FindField('ARS_RESSOURCE').asstring ;
      if Result = '' then Result:=code else Result:=Result+';'+code;
      end;
    L.ClearSelected;
    end;
FiniMove;
end;

procedure TOF_RESSOURCE_MUL.ConstituePlusTypeRessource (TheTypes : string);
var lesTypes,unType,TheSelection,TheWhere : string;
begin

  if TheTypes = '' then exit;

  LesTypes := TheTypes;
  TheSelection := '';
  TheWhere := '';
  Untype := READTOKENPipe (LesTypes,',');

  repeat
    if copy(UnType,1,1)='-' then
    begin
      if TheSelection = '' then
         TheSelection := TheSelection + 'AND ((CC_CODE <> "'+copy(UnType,2,255)+'")'
      else
        TheSelection := TheSelection + ' (CC_CODE <> "'+copy(UnType,2,255)+'")';
      if TheWhere = '' then
      begin
        TheWhere := '(ARS_TYPERESSOURCE <> "'+copy(UnType,2,255)+'")'
      end else
      begin
        TheWhere := TheWhere + ' AND (ARS_TYPERESSOURCE <> "'+copy(UnType,2,255)+'")'
      end;
    end else
    begin
      if TheSelection = '' then
        Begin
        TheSelection := TheSelection + 'AND ((CC_CODE = "'+UnType+'")';
        TheTypes := UnType;
        if Ecran.Name = 'BTRESSRECH_MUL' then
	         TypeMultiRess.Enabled := False
        else
	         TypeRessource.Enabled := False;
        end
      else
        Begin
        TheSelection := TheSelection + ' OR (CC_CODE = "'+UnType+'")';
        TheTypes := TheTypes + ';' + UnType;
        if Ecran.Name = 'BTRESSRECH_MUL' then
	         TypeMultiRess.Enabled := True
        else
	         TypeRessource.Enabled := True;
        end;
      if TheWhere = '' then
         begin
         TheWhere := '(ARS_TYPERESSOURCE = "'+ UnType +'")'
         end
      else
         begin
         TheWhere := TheWhere + ' OR (ARS_TYPERESSOURCE = "'+ UnType +'")'
         end;
    end;
    Untype := READTOKENPipe (LesTypes,',');
  until UnType = '';

  if TheSelection <> '' then
     begin
     TheSelection := TheSelection + ')';
     if Ecran.Name = 'BTRESSRECH_MUL' then
        Begin
        TypeMultiRess.Plus :=  TheSelection;
        TypeMultiRess.Value := TheTypes;
        end
     else
        Begin
        TypeRessource.Plus :=  TheSelection;
        TypeRessource.Value := TheTypes;
        end;
     end;

  if TheWhere <> '' then THEdit(GetControl('XX_WHERE')).text := TheWhere;

end;

procedure TOF_RESSOURCE_MUL.AfterShow;
var
    youser : string;
begin
//{$IFNDEF EAGLCLIENT} // attente version 560h
  Tfmul(Ecran).ForceSelectFiltre(StFiltre+';'+youser,true);
//{$ENDIF}
end ;

Procedure AFLanceFiche_Mul_Ressource;
begin
{$IFDEF BTP}
AGLLanceFiche('BTP','BTRESSOURCE_Mul','','','');
{$ELSE}
AGLLanceFiche('AFF','RESSOURCE_Mul','','','');
{$ENDIF}
end;

Procedure AFLanceFiche_Mul_RessourceLine(Argument : String);
begin
{$IFDEF BTP}
AGLLanceFiche('BTP','BTRESSOURCE_MUL','','',Argument);
{$ELSE}
AGLLanceFiche('AFF','RESSOURCE_Mul','','',Argument);
{$ENDIF}
end;

Function AFLanceFiche_Rech_Ressource(Range,Argument:string):variant;
begin
{$IFDEF BTP}
result:=AGLLanceFiche('BTP','BTRESSRECH_Mul',Range,'',Argument);
//result:=AGLLanceFiche('BTP','BTRESSOURCE_Mul',Range,'',Argument);
{$ELSE}
result:=AGLLanceFiche('AFF','RESSOURCERECH_Mul',Range,'',Argument);
{$ENDIF}
end;

procedure TOF_RESSOURCE_MUL.FlisteDblClick(Sender: TObject);
Var StArg 		: String;
    Ressource : String;
    UnType		: String;
begin

  //Si pas d'enreg double click interdit ==> FV 05112008
  //TypeRessource := GetControlText('ARS_TYPERESSOURCE'); //Fliste.DataSource.DataSet.FindField('ARS_TYPERESSOURCE').AsString;
	if FListe.datasource.DataSet.RecordCount = 0  then exit;

  if (Ecran.Name = 'BTRESSRECH_MUL') or (ecran.name = 'RTACTIONS') then
  begin
  	TFMul(Ecran).Retour :=Fliste.DataSource.DataSet.FindField('ARS_RESSOURCE').AsString ;
  	TFMUL(Ecran).Close;
    exit;
  end else
  begin
	   UnType := TypeRessource.Value;
  end;

  Ressource	:= Fliste.DataSource.DataSet.FindField('ARS_RESSOURCE').AsString;

  if Action='RECH' then
     Begin
     TFMul(Ecran).Retour := Ressource;
     TFMUL(Ecran).Close;
     End
  else
  Begin
    OpenForm.Resource(Ressource, UnType);
(*
    stArg:='ACTION=MODIFICATION;';
    if Not AGLJaiLeDroitFiche(['RESSOURCE',StArg],2) Then exit;
      Begin
      //Aiguillage pour gestion écran Ressource ou Matériel
      //Non en line
      AGLLanceFiche ('BTP','BTRESSOURCE','',Ressource, 'TYPERESSOURCE=' + UnType);
      end;
*)
  end;

     if GetParamSocSecur('SO_REFRESHMUL', true) then refreshdb;


end;

procedure TOF_RESSOURCE_MUL.CreateSalarie(Sender: TObject);
begin

	CreationRessource('SAL');

end;

procedure TOF_RESSOURCE_MUL.CreateAutre(Sender: TObject);
begin

	CreationRessource('AUT');

end;

procedure TOF_RESSOURCE_MUL.CreateIterim(Sender: TObject);
begin

	CreationRessource('INT');

end;

procedure TOF_RESSOURCE_MUL.CreateLocation(Sender: TObject);
begin

	CreationRessource('LOC');

end;

procedure TOF_RESSOURCE_MUL.CreateMateriel(Sender: TObject);
begin

	CreationRessource('MAT');

end;

procedure TOF_RESSOURCE_MUL.CreateOutil(Sender: TObject);
begin

	CreationRessource('OUT');

end;

procedure TOF_RESSOURCE_MUL.CreateSsTraitance(Sender: TObject);
begin

	CreationRessource('ST');

end;

procedure TOF_RESSOURCE_MUL.BInsert_OnClick(Sender: TObject);
begin

  if GetControlText('ARS_TYPERESSOURCE') <> '' then
	   CreationRessource(GetControlText('ARS_TYPERESSOURCE'));

end;

Procedure TOF_RESSOURCE_MUL.CreationRessource(TypeRessource : String);
Begin

		if not (Ecran is TFMul) then exit;

		if TypeRessource <> '' then
       V_PGI.DispatchTT(6,taCreat,'','','TYPERESSOURCE='+ TypeRessource)
    Else
       V_PGI.DispatchTT(6,taCreat,'','','');

  	if GetParamSocSecur('SO_REFRESHMUL', true) then TToolBarButton97(GetCOntrol('Bcherche')).Click;

end;

Initialization
registerclasses([TOF_RESSOURCE_MUL]);
//grc
RegisterAglFunc('RameneListeRessources', TRUE , 0, AGLRameneListeRessources);
end.


