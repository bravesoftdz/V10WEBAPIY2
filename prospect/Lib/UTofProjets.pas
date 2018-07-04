unit UTofProjets;

interface
uses  Classes,
      HCtrls,HEnt1,UTOF,UtilRT,ParamSoc,UtilSelection,forms,HTB97,HMsgBox,
{$ifdef AFFAIRE}
UTOFAFTRADUCCHAMPLIBRE,
{$endif}
{$IFDEF EAGLCLIENT}
      MaineAGL,eMul;
{$ELSE}
      HDB,Fe_Main,Mul;
{$ENDIF}
Type
{$ifdef AFFAIRE}
                //PL le 18/05/07 pour gérer les champs libres si paramétrés
     TOF_Projets = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
     TOF_Projets = Class (TOF)
{$endif}
     Private
        VerrouModif :boolean;
        StFiltre : string ;
{$IFDEF EAGLCLIENT}
        Fliste : THGrid;
{$ELSE}
        Fliste : THDbGrid;
{$ENDIF}

        procedure AfterShow;
    		procedure FListe_OnDblClick(Sender: TObject);
    		procedure Binsert_OnClick(Sender: TObject);
     public
        procedure OnArgument(Arguments : String ) ; override ;
        procedure OnLoad ; override ;
     END ;

Function RTLanceFiche_Projets(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

implementation

procedure TOF_Projets.OnArgument(Arguments : String ) ;
var x:integer;
    Critere,ChampMul,ValMul,stArg : String;
begin
inherited ;
VerrouModif := False;
x := pos('CONSULTATION',Arguments);
if x <> 0 then VerrouModif :=true;
  stArg:=Arguments;
  Critere:=(ReadTokenSt(stArg));
  repeat
      Critere:=(ReadTokenSt(stArg));
      if Critere <> '' then
        begin
        x:=pos('=',Critere);
        if x<>0 then
           begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));
           if ChampMul='FILTRE' then StFiltre := ValMul
           end;
        end;
  until Critere='';
  if Ecran.Name='RTPROJETS_TL' then
    begin
    Tfmul(Ecran).FiltreDisabled:=true;
    TFmul(Ecran).OnAfterFormShow := AfterShow;
    end;

  if GetParamSocSecur('SO_RTGESTINFOS00Q',False) then
      MulCreerPagesCL(TForm (Ecran),'NOMFIC=RTPROJETS');
{$IFDEF EAGLCLIENT}
  Fliste := THGrid(GetControl('Fliste'));
{$ELSE}
  Fliste := THdbGrid(GetControl('Fliste'));
{$ENDIF}

  Fliste.OnDblClick := FListe_OnDblClick;
  TToolBarButton97(GetControl('BINSERT')).OnClick := Binsert_OnClick;
end;

procedure TOF_Projets.OnLoad;
begin
  inherited;
  if (TFMul(Ecran).name = 'RTPROJETS_MUL') then
  begin
  SetControlText('XX_WHERE',RTXXWhereConfident('CON')) ;
  end;

  if (TFMul(Ecran).name = 'RTPROJETS_TIERS')  then
  begin
    SetControlChecked('BVERROU',VerrouModif);
    SetControlEnabled('BINSERT',not VerrouModif);
  end;
end;

Function RTLanceFiche_Projets(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_Projets.AfterShow;
begin
  Tfmul(Ecran).ForceSelectFiltre(StFiltre , V_PGI.User,false,true);
end ;

procedure TOF_Projets.FListe_OnDblClick(Sender : TObject);
var stArg,Projet,Tiers : string;

begin
{$IFDEF EAGLCLIENT}
	if FListe.RowCount < 1 then exit;
	TFMul(ecran).Q.TQ.Seek(FListe.Row-1) ;

	Projet:=TFMul(ecran).Q.FindField('RPJ_PROJET').AsString;
	Tiers:=TFMul(ecran).Q.FindField('RPJ_TIERS').AsString;
{$ELSE}
	if FListe.datasource.DataSet.RecordCount = 0  then exit;

	Projet:=Fliste.datasource.dataset.FindField('RPJ_PROJET').AsString;
	Tiers:=Fliste.datasource.dataset.FindField('RPJ_TIERS').AsString;
{$ENDIF}


if Projet <> '' then
  begin
  stArg:='ACTION=MODIFICATION';
  if (Tiers <> '') AND (RTDroitModiftiers(Tiers)=False) then stArg:= 'ACTION=CONSULTATION';
  AglLanceFiche('RT','RTPROJETS','',Projet,StArg+';MONOFICHE') ;
  TtoolBarButton97(GetCOntrol('Bcherche')).Click;
  end;
end;

procedure TOF_Projets.Binsert_OnClick(Sender : TObject);
begin
  if AGLRTExisteConfident([],0) <> '' then
  begin
    AglLanceFiche('RT','RTPROJETS','','','ACTION=CREATION') ;
    TtoolBarButton97(GetCOntrol('Bcherche')).Click;
  end
  else
  	PGiBox ('Vous n''avez pas les droits d''accès en création');
end;


Initialization
registerclasses([TOF_Projets]);
end.
