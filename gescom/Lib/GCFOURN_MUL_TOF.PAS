{***********UNITE*************************************************
Auteur  ...... : Joël TRIFILIEFF
Créé le ...... : 25/04/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : GCFOURN_MUL ()
Mots clefs ... : TOF;GCFOURN_MUL
*****************************************************************}
Unit GCFOURN_MUL_TOF ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$ELSE}
     maineagl,
{$ENDIF}
		HDB,
{$ifdef AFFAIRE}
      UtofAfTraducChampLibre,
{$ENDIF}
    
    forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,AGLInit,uTOFComm,mul ;

function GCLanceFiche_GcFournMul(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
  TOF_GCFOURN_MUL = Class (TOF_AFTRADUCCHAMPLIBRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (Arguments : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure GSDblClick (Sender : Tobject);
  private
    modeSelect : boolean;
  end ;

Implementation

function GCLanceFiche_GcFournMul(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
Result:='';
if Nat='' then exit;
if Cod='' then exit;
Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_GCFOURN_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GCFOURN_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GCFOURN_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_GCFOURN_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_GCFOURN_MUL.OnArgument (Arguments : String ) ;
var stArg, Critere, ChampMul : string;
    PosEgal : integer;
begin
	fMulDeTraitement := true;
  Inherited ;
  modeSelect := false;
  fTableName := 'TIERS';
// Si on ouvre le MUL en passant en paramètre ACTION=CONSULTATION (par le MDisp)
// on affecte la valeur du paramètre au champs TYPEACTION (THLabel de la fiche MUL)
// et on cache le bouton "bInsert"
	SetControlText('TYPEACTION','ACTION=MODIFICATION'); //Initialisation à modification
	stArg:=Arguments;
	Repeat
    Critere := UpperCase(Trim(ReadTokenSt(stArg)));
    if Critere<>'' then
		begin
      PosEgal := pos('=',Critere);
      if PosEgal<>0 then ChampMul := copy(Critere,1,PosEgal-1);
      if ChampMul='ACTION' then SetControlText('TYPEACTION',Critere);
      if ChampMul='SELECTION' then modeSelect := true;
    end;
  until Critere='';
	SetControlVisible('bInsert',StringToAction(GetControlText('TYPEACTION'))<>taConsult);
  THDbGrid (GetControl('FLISTE')).OnDblClick := GSDblClick;
end ;

procedure TOF_GCFOURN_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_GCFOURN_MUL.GSDblClick(Sender: Tobject);
var Auxiliaire,NatureAuxi,Tiers : string;
begin
{$IFDEF EAGLCLIENT}
	if FListe.RowCount = 0  then exit;
	TFMul(ecran).Q.TQ.Seek(TFMul(ecran).FListe.Row-1) ;
  tiers:=TFMul(ecran).Q.FindField('T_TIERS').AsString;
  Auxiliaire:=TFMul(ecran).Q.FindField('T_AUXILIAIRE').AsString;
  NatureAuxi:=TFMul(ecran).Q.FindField('T_NATUREAUXI').AsString;
{$ELSE}
	if THDbGrid(getcontrol('FListe')).datasource.DataSet.RecordCount = 0  then exit;
  Tiers:=THDbGrid(getcontrol('FListe')).datasource.dataset.FindField('T_TIERS').AsString;
  Auxiliaire:=THDbGrid(getcontrol('FListe')).datasource.dataset.FindField('T_AUXILIAIRE').AsString;
  NatureAuxi:=THDbGrid(getcontrol('FListe')).datasource.dataset.FindField('T_NATUREAUXI').AsString;
{$ENDIF}

  if modeSelect then
	begin
		//Retour=GetChamp('T_TIERS');
		TFMul(Ecran).Retour := tiers;
  	TFMul(Ecran).Close;
	end else
  begin
		AglLanceFiche('GC','GCFOURNISSEUR','', Auxiliaire,'MONOFICHE;T_NATUREAUXI=' + NatureAuxi+';'+THEdit(GetCONTROL('TYPEACTION')).text);
		refreshDB;
	end;

end;

Initialization
  registerclasses ( [ TOF_GCFOURN_MUL ] ) ;
end.
