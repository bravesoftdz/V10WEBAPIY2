{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 30/10/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CHAINAGES ()
Mots clefs ... : TOF;CHAINAGES
*****************************************************************}
Unit UtofTypesChainages ;

Interface

Uses Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     FE_Main,db,
{$ELSE}
     Maineagl,Utob,
{$ENDIF}
     sysutils,
     UTOF,
     uTableFiltre,
     SaisieList,
     hTB97

     ;

Type
  TOF_TYPESCHAINAGES = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure DoClick( Sender: TObject );
    procedure DoDblClick( Sender: TObject );
    procedure DoSetNavigate( Sender: TObject );
  Private
    stProduitpgi : string;
    TF: TTableFiltre;
  end ;

Function RTLanceFiche_TypesChainages(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

implementation

uses
  HCtrls
  ;

Function RTLanceFiche_TypesChainages(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_TYPESCHAINAGES.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_TYPESCHAINAGES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_TYPESCHAINAGES.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_TYPESCHAINAGES.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_TYPESCHAINAGES.OnArgument (S : String ) ;
begin
  Inherited ;
  stProduitpgi:=S;
  //SetControlText ('RPG_PRODUITPGI',stproduitpgi);
  if (Ecran<>nil) and (Ecran is TFSaisieList ) then
  begin
    TF := TFSaisieList(Ecran).LeFiltre;
    TF.OnSetNavigate := DoSetNavigate;
    if stProduitpgi='GRF' then
      begin
      TF.LaTreeListe:='RFPARCHAINAGES';
      TF.LaGridListe:='RFPARACTIONS';
      SetControlText('RPG_PRODUITPGI','GRF') ;
      end
    else
      SetControlText('RPG_PRODUITPGI','GRC') ;
  end;
  THTreeView(GetControl( 'TreeEntete' )).OnDblClick := DoDblClick;
  TToolBarButton97(GetControl( 'BNEWCHAINAGE' )).OnClick := DoClick;

end ;

procedure TOF_TYPESCHAINAGES.DoSetNavigate( Sender: TObject );
begin
if (TF.TOBFiltre.detail.count = 0) or ( TF.State = dsInsert) then SetControlEnabled('BZOOM',false)
else SetControlEnabled('BZOOM',true);
if (TF.TOBFiltre.detail.count = 0) then
  begin
  SetControlEnabled('BFLECHEBAS',False);
  SetControlEnabled('BFLECHEHAUT',False);
  end;
end;

procedure TOF_TYPESCHAINAGES.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_TYPESCHAINAGES.DoClick( Sender: TObject );
var
  CodeChainage : String;
begin
  if UpperCase(TControl( Sender ).name) = 'BNEWCHAINAGE' then
    begin
    CodeChainage := AGLLanceFiche('RT','RTTYPECHAINAGE','','','MONOFICHE;ACTION=CREATION;PRODUITPGI='+stProduitpgi) ;
    if CodeChainage <> '' then
      TF.RefreshEntete( CodeChainage+';'+stProduitpgi);
    end;
end;

procedure TOF_TYPESCHAINAGES.DoDblClick( Sender: TObject );
var
  CodeChainage : String;
begin
  if UpperCase(TControl( Sender ).name) = 'TREEENTETE' then
    begin
    CodeChainage := AGLLanceFiche('RT','RTTYPECHAINAGE','',TF.TOBFiltre.GetValue('RPG_CHAINAGE')+';'+stProduitpgi,'ACTION=MODIFICATION;MONOFICHE;PRODUITPGI='+stProduitpgi);
    if CodeChainage <> '' then
      TF.RefreshEntete( CodeChainage+';'+stProduitpgi );
    end;
end;



Initialization
  registerclasses ( [ TOF_TYPESCHAINAGES ] ) ;
end.

