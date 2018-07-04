{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 14/10/2008
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PROFILART_MUL ()
Mots clefs ... : TOF;PROFILART_MUL
*****************************************************************}
Unit ProfileArt_Mul ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS}
     dbtables,
     {$ELSE}
     uDbxDataSet,
     {$ENDIF}
     mul,
     DBGrids,
     FE_Main,
{$else}
     eMul,
     Maineagl,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Ent1,
     HMsgBox,
     //
     HTB97,
     HDB,
     ParamSoc,
     //
     UTOF ; 

Type
  TOF_PROFILART_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  //définition particulière à la fiche
  Private
    BInsert : TToolBarButton97;
    BProfil	: TToolBarButton97;
  //
    Fliste 	: THDbGrid;
  //
  	GPF_FAMILLENIV1 : THValComboBox;
	  GPF_FAMILLENIV2 : THValComboBox;
    GPF_FAMILLENIV3 : THValComboBox;
  //
    ProfileArticle	: String;
    
  //Modif FV
		Procedure BInsert_OnClick(Sender: TObject);
    Procedure BProfil_OnClick(Sender: TObject);
    procedure DefinitionEcran;
	  Procedure FlisteDblClick (Sender : TObject);
    procedure GPF_FAMILLENIV1_OnChange(Sender: TObject);
    procedure GPF_FAMILLENIV2_OnChange(Sender: TObject);

  Public

  end ;

Implementation

procedure TOF_PROFILART_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PROFILART_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PROFILART_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PROFILART_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PROFILART_MUL.OnArgument (S : String ) ;
var CC      : THValComboBox;
begin
  Inherited ;

  //Gestion des zones écrans
  DefinitionEcran;

  
  //Gestion Restriction Domaine
  CC:=THValComboBox(GetControl('GPF_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

end;

procedure TOF_PROFILART_MUL.DefinitionEcran;
Var HLabel : THLabel;
begin

  HLabel := THLabel(GetControl('TGPF_FAMILLENIV1'));
	HLabel.Caption := RechDom('GCLIBFAMILLE','LF1', True);

  HLabel := THLabel(GetControl('TGPF_FAMILLENIV2'));
	HLabel.Caption := RechDom('GCLIBFAMILLE','LF2', True);

  HLabel := THLabel(GetControl('TGPF_FAMILLENIV3'));
	HLabel.Caption := RechDom('GCLIBFAMILLE','LF3', True);

  Fliste := THDbGrid (GetControl('FLISTE'));
  Fliste.OnDblClick := FlisteDblClick;

  BInsert := TToolbarButton97(ecran.FindComponent('BInsert'));
  BInsert.OnClick := BInsert_OnClick;

  BProfil := TToolbarButton97(ecran.FindComponent('BProfil'));
  BProfil.OnClick := BProfil_OnClick;

 	GPF_FAMILLENIV1 := THValComboBox(GetControl('GPF_FAMILLENIV1'));
  GPF_FAMILLENIV1.OnChange := GPF_FAMILLENIV1_Onchange;
  GPF_FAMILLENIV2 := THValComboBox(GetControl('GPF_FAMILLENIV2'));
  GPF_FAMILLENIV2.Onchange := GPF_FAMILLENIV2_Onchange;
  GPF_FAMILLENIV3 := THValComboBox(GetControl('GPF_FAMILLENIV3'));

end ;

procedure TOF_PROFILART_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PROFILART_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PROFILART_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PROFILART_MUL.BInsert_OnClick(Sender: TObject);
begin

//uniquement en line
//AGLLanceFiche('BTP','BTPROFILART_S1','','','ACTION=CREATION');

	AGLLanceFiche('BTP','BTPROFILART','','','ACTION=CREATION');

	TToolBarButton97(GetControl('Bcherche')).Click;

end;

procedure TOF_PROFILART_MUL.FlisteDblClick(Sender: TObject);
begin

	ProfileArticle := Fliste.datasource.dataset.FindField('GPF_PROFILARTICLE').AsString;

  TFMul(Ecran).Retour := ProfileArticle;
  TFMUL(Ecran).Close;

end;

procedure TOF_PROFILART_MUL.BProfil_OnClick(Sender: TObject);
begin

	ProfileArticle := Fliste.datasource.dataset.FindField('GPF_PROFILARTICLE').AsString;

//uniquement en line
//	AGLLanceFiche('BTP','BTPROFILART_S1','',ProfileArticle,'');

	AGLLanceFiche('BTP','BTPROFILART','',ProfileArticle,'');

	TToolBarButton97(GetControl('Bcherche')).Click;

end;

Procedure TOF_PROFILART_MUL.GPF_FAMILLENIV1_OnChange(Sender: TObject);
Begin

	if GetParamSoc('SO_GCFAMHIERARCHIQUE') then
 		 begin
	   GPF_FAMILLENIV2.Plus := '';
	   GPF_FAMILLENIV3.Plus := '';
	   if (GPF_FAMILLENIV1.Value <> '') then
    	  begin
		    GPF_FAMILLENIV2.Plus := 'and CC_LIBRE like "%'+GPF_FAMILLENIV1.Value+'%"';
    		GPF_FAMILLENIV3.Plus := 'and CC_LIBRE like "%'+GPF_FAMILLENIV1.Value+'%"';
		    end;
		 end;

end;

Procedure TOF_PROFILART_MUL.GPF_FAMILLENIV2_OnChange(Sender: TObject);
var st : String;
Begin

	if GetParamSoc('SO_GCFAMHIERARCHIQUE') then
		 begin
  	 st := '';
	   GPF_FAMILLENIV3.Plus := '';
	   if GPF_FAMILLENIV1.Value <> '' then st := st+GPF_FAMILLENIV1.Value;
	   if GPF_FAMILLENIV2.Value <> '' then st := st+GPF_FAMILLENIV2.Value;
	   if st <> '' then GPF_FAMILLENIV3.Plus := 'and CC_LIBRE like "%'+st+'%"';
  	 end;

end;


Initialization
  registerclasses ( [ TOF_PROFILART_MUL ] ) ;
end.

