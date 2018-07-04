{***********UNITE*************************************************
Auteur  ...... : SANTUCCI Lionel
Créé le ...... : 20/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PARAMCATTAXE ()
Mots clefs ... : TOF;PARAMCATTAXE
*****************************************************************}
Unit PARAMCATTAXE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     uTob,
     MainEagl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Graphics,
     Grids,
     HTB97,
     UTOB,
     Vierge,
     Ent1,
     EntGc,
     AglInit,
     UTOF ;

Type
  TOF_PARAMCATTAXE = Class (TOF)
      procedure OnNew                    ; override ;
      procedure OnDelete                 ; override ;
      procedure OnUpdate                 ; override ;
      procedure OnLoad                   ; override ;
      procedure OnArgument (S : String ) ; override ;
      procedure OnDisplay                ; override ;
      procedure OnClose                  ; override ;
      procedure OnCancel                 ; override ;
    private
   		GS : THGrid;
      TOBC : TOB;
      BPT_CATEGORIETAXE,BPT_LIBELLE : THEdit;
      BPT_TYPETAXE : THValComboBox;
      BACCEPTMODIF,BANNULMODIF,Bferme,Bvalider : TToolbarButton97;
      procedure SetEnabled (Actif : boolean);
      procedure SetEvent;
			procedure GSEnter (Sender : TObject);
    	procedure GSDblClick(Sender: TObject);
      procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
      procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
  		procedure ChampsExit (Sender : TObject);
  		procedure BACCEPTMODIFClick (Sender :TObject);
  		procedure BANNULMODIFClick (Sender :TObject);
      procedure AffichelaLigne(LaLigne : integer);
      procedure EnregistreModif;
      procedure MiseAjourTablette;
    procedure AfficheLaGrille;
    function ModifAutorise: boolean;
    procedure InitTOB;
  end ;

procedure ParametrageTypeTaxe ;

Implementation

procedure ParametrageTypeTaxe;
begin
	Thetob := VH_GC.TOBParamTaxe;
	AGLLanceFiche ('BTP','BTPARAMCATTAXE','','','ACTION=MODIFICATION');
end;

procedure TOF_PARAMCATTAXE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PARAMCATTAXE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PARAMCATTAXE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PARAMCATTAXE.OnLoad ;
begin
  Inherited ;
end ;


procedure TOF_PARAMCATTAXE.InitTOB;
var Indice : integer;
		LaTOB : TOB;
begin
	for Indice := 0 to VH_GC.TOBParamTaxe.detail.count -1 do
  begin
  	LaTOB := VH_GC.TOBParamTaxe.detail[Indice];
    if not LaTob.FieldExists ('_MODIFICATION') then
    begin
      LaTOB.AddChampSupValeur ('_MODIFICATION','-',True);
    end else
    begin
      LaTOB.PutValue ('_MODIFICATION','-');
    end;
	end;
end;

procedure TOF_PARAMCATTAXE.AfficheLaGrille;
var Indice : integer;
begin
	GS.ColLengths [0] := 20;
	GS.ColLengths [1] := 80;
	for Indice := 0 to VH_GC.TOBParamTaxe.detail.count -1 do
  begin
  	AfficheLaLigne (Indice+1);
  end;
end;

procedure TOF_PARAMCATTAXE.OnArgument (S : String ) ;
var Cancel : boolean;
begin
  Inherited ;
  Cancel := false;
  GS := THGrid(GetCOntrol('GS'));
  BPT_CATEGORIETAXE := THEdit(GetControl('BPT_CATEGORIETAXE'));
  BPT_LIBELLE := THEdit(GetControl('BPT_LIBELLE'));
  BPT_TYPETAXE := THValComboBox (GetCOntrol('BPT_TYPETAXE'));
  BFerme := TToolbarButton97  (GetCOntrol('Bferme'));
  BValider := TToolbarButton97  (GetCOntrol('BValider'));
  BACCEPTMODIF := TToolbarButton97  (GetCOntrol('BACCEPTMODIF'));
	BANNULMODIF := TToolbarButton97  (GetCOntrol('BANNULMODIF'));
  GS.rowCount := VH_GC.TOBParamTaxe.detail.count+1;
  InitTOB;
  AfficheLaGrille;
  GS.row := 1;
  GSRowEnter (self,GS.row,cancel,false);
  SetEnabled (false);
  SetEvent;
  TFVierge(ecran).HMTrad.ResizeGridColumns (GS);
end ;

procedure TOF_PARAMCATTAXE.OnClose ;
begin
  Inherited ;
	AvertirTable ('GCCATEGORIETAXE');
end ;

procedure TOF_PARAMCATTAXE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PARAMCATTAXE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PARAMCATTAXE.SetEnabled(Actif: boolean);
begin
  BPT_LIBELLE.Enabled := Actif;
  BPT_TYPETAXE.Enabled := Actif;
  //
  BACCEPTMODIF.Visible := Actif;
  BANNULMODIF.Visible := Actif;
  //
  BFerme.Visible := (not Actif);
  BValider.Visible := (not Actif);
  GS.enabled := (not Actif);
  //
  if Actif then
  begin
    BPT_LIBELLE.Color := clWindow;
    BPT_TYPETAXE.Color := clWindow;
  end else
  begin
    BPT_LIBELLE.Color := clInactiveCaption;
    BPT_TYPETAXE.Color := clInactiveCaption;
  end;
end;

procedure TOF_PARAMCATTAXE.SetEvent;
begin
	GS.OnEnter := GSEnter;
	GS.OnRowEnter := GSRowEnter;
	GS.OnRowExit := GSRowExit;
	GS.OnDblClick  := GSDblClick;
  BPT_LIBELLE.OnExit := ChampsExit;
  BPT_TYPETAXE.OnExit := ChampsExit;
  BACCEPTMODIF.OnClick := BACCEPTMODIFClick;
  BANNULMODIF.OnClick := BANNULMODIFClick;
end;

procedure TOF_PARAMCATTAXE.BACCEPTMODIFClick(Sender: TObject);
begin
	EnregistreModif;
	SetEnabled(false);
end;

procedure TOF_PARAMCATTAXE.ChampsExit(Sender: TObject);
var NewValue,NOM : string;
begin
	if Sender is THEdit Then
  begin
  	NewValue := THEdit(Sender).Text;
    Nom := THEDit(Sender).Name;
  end else if Sender is THValComboBox then
  begin                                                                          	
    NewValue := THValComboBox(Sender).Value;
    Nom := THValComboBox(Sender).Name;
  end;
  if TOBC.GetValue (Nom) <> NewValue then TOBC.PutValue('_MODIFICATION','X');
end;

procedure TOF_PARAMCATTAXE.GSEnter(Sender: TObject);
var cancel : boolean;
begin
  cancel := false;
  GSRowEnter(self,1,cancel,false);
end;

procedure TOF_PARAMCATTAXE.GSRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin
  TOBC := VH_GC.TOBParamTaxe.detail[GS.row-1];
  TOBC.PutEcran (ecran);
end;

procedure TOF_PARAMCATTAXE.GSRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
var laTOB,TOBSAV : TOB;
begin
	LaTOB := VH_GC.TOBParamTaxe.detail[Ou-1];
  if (laTOB.getValue('_MODIFICATION')= 'X') then
  begin
  	if PGIAsk (TraduireMemoire('Voules-vous enregistrer ces modifications')) = mryes then
    begin
      TOBSAV := TOBC;
      TOBC := LaTOB;
      EnregistreModif;
      TOBC := TOBSAV;
    end;
  end;
  SetEnabled (false);
end;

function TOF_PARAMCATTAXE.ModifAutorise : boolean;
var LaTOB,TOBCUr : TOB;
		LeCode : string;
begin
	result := true;
  if GS.row > 1 then
  begin
    LaTOB := VH_GC.TOBParamTaxe.detail[GS.row-2];
    if LaTOB.getValue('BPT_TYPETAXE')='' Then BEGIN result := false; Exit; END;
  end;
  TOBCur := VH_GC.TOBParamTaxe.detail[GS.row-1];
  LeCode := TOBCur.GetValue('BPT_CATEGORIETAXE');
  if (VH^.DefCatTVA=leCode) or (VH^.DefCatTPF=LeCode) then result := false;
end;

procedure TOF_PARAMCATTAXE.GSDblClick(Sender: TObject);
var IsModifAutorise : boolean;
begin
  IsModifAutorise := ModifAutorise;
	if ModifAutorise then SetEnabled (True);
end;

procedure TOF_PARAMCATTAXE.BANNULMODIFClick(Sender: TObject);
var Cancel : boolean;
begin
	Cancel := false;
  TOBC := VH_GC.TOBParamTaxe.Detail[GS.row-1];
  AffichelaLigne(GS.row);
  GSRowEnter (self,GS.row,cancel,false);
  TOBC.PutValue('_MODIFICATION','-');
	SetEnabled(false);
end;

procedure TOF_PARAMCATTAXE.AffichelaLigne(LaLigne: integer);
var TOBC : TOB;
begin
  TOBC := VH_GC.TOBParamTaxe.detail[LaLigne-1];
  TOBC.PutLigneGrid (GS,LaLigne,false,false,'BPT_CATEGORIETAXE;BPT_LIBELLE');
end;

procedure TOF_PARAMCATTAXE.EnregistreModif;
begin
	TOBC.GetEcran (ecran);
  BEGINTRANS;
  TRY
    if not TOBC.UpdateDB then V_PGI.IOError := OEUnknown;
    MiseAjourTablette;
    AffichelaLigne (GS.row);
  	COMMITTRANS;
  	TOBC.PutValue('_MODIFICATION','-');
  EXCEPT
  	PGIError (TraduireMemoire('Erreur durant la mise à jour'));
    GetParamTaxe; // on recharge
  	TOBC.PutValue('_MODIFICATION','-');
  	ROLLBACK;
  END;
end;

procedure TOF_PARAMCATTAXE.MiseAjourTablette;
var TOBCC : TOB;
begin
	TOBCC := TOB.Create ('CHOIXCOD',nil,-1);
  TRY
    TOBCC.PutValue('CC_TYPE','GCX');
    TOBCC.PutValue('CC_CODE',TOBC.GetValue('BPT_CATEGORIETAXE'));
    TOBCC.PutValue('CC_LIBELLE',TOBC.GetValue('BPT_LIBELLE'));
    TOBCC.PutValue('CC_ABREGE',Copy(TOBC.GetValue('BPT_LIBELLE'),1,17));
    if not TOBCC.UpdateDB (false) Then V_PGI.IOError := oeUnknown;
  FINALLY
  TOBCC.free;
  END;
end;

Initialization
  registerclasses ( [ TOF_PARAMCATTAXE ] ) ;
end.

