{***********UNITE*************************************************
Auteur  ...... : VAUTRAIN Franck (avec l'aide de Lionel)
Créé le ...... : 23/02/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTRECONDUCTION ()
Mots clefs ... : TOF;RECONDUCTION
*****************************************************************}
Unit RECONDUCTION_TOF ;

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
  TOF_RECONDUCTION = Class (TOF)
      procedure OnNew                    ; override ;
      procedure OnDelete                 ; override ;
      procedure OnUpdate                 ; override ;
      procedure OnLoad                   ; override ;
      procedure OnArgument (S : String ) ; override ;
      procedure OnDisplay                ; override ;
      procedure OnClose                  ; override ;
      procedure OnCancel                 ; override ;

    private
   	  GS   : THGrid;
      TOBC : TOB;
      Ok_Creation : Boolean;
      BRE_CODE,BRE_LIBELLE : THEdit;
      BRE_TYPEACTION : THValComboBox;
		  BRE_GENERE : TCheckBox;
      BRE_NBMOIS : THSpinEdit;
      BACCEPTMODIF,BANNULMODIF,Binsert, Bdelete, Bferme,Bvalider : TToolbarButton97;
      TOBReconduction : TOB;
      procedure SetEnabled (Actif : boolean);
      procedure SetEvent;
      procedure GSEnter (Sender : TObject);
      procedure GSDblClick(Sender: TObject);
      procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
      procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
  	  procedure ChampsExit (Sender : TObject);
  	  procedure BACCEPTMODIFClick (Sender :TObject);
  	  procedure BANNULMODIFClick (Sender :TObject);
	    procedure BInsertClick (Sender :TObject);
  	  procedure BDeleteClick (Sender :TObject);
      procedure AffichelaLigne(LaLigne : integer);
      procedure EnregistreModif;
      procedure MiseAjourTablette;
      procedure AfficheLaGrille;
      function ModifAutorise: boolean;
      Function ControleCodeReconduction(code : string) : TOB;
      Function ControleAffaire : Boolean;
      procedure InitTOB;
      procedure SupprimeEnreg (TOBS : TOB);
      procedure SupprimeTablette;
  end ;

procedure ParametrageTypeAction ;

Implementation

procedure ParametrageTypeAction;
begin
	AGLLanceFiche ('BTP','BTRECONDUCTION','','','ACTION=MODIFICATION');
end;

procedure TOF_RECONDUCTION.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_RECONDUCTION.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_RECONDUCTION.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_RECONDUCTION.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_RECONDUCTION.InitTOB;
var Indice : integer;
	LaTOB : TOB;
    QQ : TQuery;
begin
  QQ := OpenSql ('SELECT * from BRECONDUCTION', True,-1,'',true);
  TOBreconduction.LoadDetailDB ('BRECONDUCTION','','',QQ,false,true);
  ferme (QQ);
  for Indice := 0 to TOBReconduction.detail.count -1 do
  begin
  	LaTOB := TOBReconduction.detail[Indice];
    LaTOB.AddChampSupValeur ('_NEW','-');
    LaTOB.AddChampSupValeur ('_MODIFICATION','-',True);
  end;
end;

procedure TOF_RECONDUCTION.AfficheLaGrille;
var Indice : integer;
begin
	GS.ColLengths [0] := 20;
	GS.ColLengths [1] := 80;
	for Indice := 0 to TOBReconduction.detail.count -1 do
  begin
  	AfficheLaLigne (Indice+1);
  end;
end;

procedure TOF_RECONDUCTION.OnArgument (S : String ) ;
var Cancel : boolean;
begin
  Inherited ;

  TOBReconduction:=TOB.Create ('LES BRECONDUCTION',nil,-1);

  Cancel := false;
  GS := THGrid(GetCOntrol('GS'));

  BRE_CODE := THEdit(GetControl('BRE_CODE'));
  BRE_LIBELLE := THEdit(GetControl('BRE_LIBELLE'));
  BRE_TYPEACTION := THValComboBox (GetCOntrol('BRE_TYPEACTION'));
  BRE_GENERE := TCheckBox(GetCONTROL('BRE_GENERE'));
  BRE_NBMOIS := THSpinEdit(GetCONTROL('BRE_NBMOIS'));

  BFerme := TToolbarButton97  (GetCOntrol('Bferme'));
  BValider := TToolbarButton97  (GetCOntrol('BValider'));
  BACCEPTMODIF := TToolbarButton97  (GetCOntrol('BACCEPTMODIF'));
  BANNULMODIF := TToolbarButton97  (GetCOntrol('BANNULMODIF'));
  Binsert := TToolbarButton97  (GetCOntrol('BInsert'));
  BDelete := TToolbarButton97  (GetCOntrol('BDelete'));

  InitTOB;
  GS.rowCount := TOBReconduction.detail.count+1;
  AfficheLaGrille;
  GS.row := 1;
  GSRowEnter (self,GS.row,cancel,false);
  SetEnabled (false);
  SetEvent;

  TFVierge(ecran).HMTrad.ResizeGridColumns (GS);

end ;

procedure TOF_RECONDUCTION.OnClose ;
begin
  Inherited ;
  TOBReconduction.free;
  AvertirTable ('TTRECONDUCABO');
end ;

procedure TOF_RECONDUCTION.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_RECONDUCTION.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_RECONDUCTION.SetEnabled(Actif: boolean);
begin

  BRE_CODE.Enabled := Actif;
  BRE_LIBELLE.Enabled := Actif;
  BRE_TYPEACTION.Enabled := Actif;
	BRE_GENERE.Enabled := Actif;
	BRE_NBMOIS.Enabled := Actif;
  //
  BACCEPTMODIF.Visible := Actif;
  BANNULMODIF.Visible := Actif;
  //
  BFerme.Visible := (not Actif);
  BValider.Visible := (not Actif);
  Binsert.visible := (not Actif);
  BDelete.visible := (not Actif);
  GS.enabled := (not Actif);
  //
  if Actif then
  begin
    BRE_CODE.Color := clWindow;
    BRE_LIBELLE.Color := clWindow;
    BRE_TYPEACTION.Color := clWindow;
	  BRE_GENERE.Color := ClBtnFace;
	  BRE_NBMOIS.Color := clWindow;
    BRE_CODE.SetFocus;
  end else
  begin
    BRE_CODE.Color := clInactiveCaptionText;
    BRE_LIBELLE.Color := clInactiveCaptionText;
    BRE_TYPEACTION.Color := clInactiveCaptionText;
	  BRE_GENERE.Color := clInactiveCaptionText;
	  BRE_NBMOIS.Color := clInactiveCaptionText;
  end;
end;

procedure TOF_RECONDUCTION.SetEvent;
begin

  GS.OnEnter := GSEnter;
  GS.OnRowEnter := GSRowEnter;
  GS.OnRowExit := GSRowExit;
  GS.OnDblClick  := GSDblClick;

  BRE_CODE.OnExit := ChampsExit;
  BRE_LIBELLE.OnExit := ChampsExit;
  BRE_TYPEACTION.OnExit := ChampsExit;
  BRE_GENERE.OnExit := ChampsExit;
  BRE_NBMOIS.OnExit := ChampsExit;

  BACCEPTMODIF.OnClick := BACCEPTMODIFClick;
  BANNULMODIF.OnClick := BANNULMODIFClick;
  Binsert.OnClick := BinsertClick;
  BDelete.OnClick := BDeleteClick;

end;

Procedure TOF_RECONDUCTION.BDeleteClick(Sender: TObject);
var cancel : boolean;
begin
  cancel := false;

  If not ControleAffaire then
  Begin
     PGIBox(traduirememoire('Suppression Impossible, code reconduction présent sur une affaire'), ecran.caption);
     exit;
  end;

  TOBC := TObreconduction.Detail[gs.row -1];
  SupprimeEnreg (TOBC);
  SupprimeTablette;
  TOBC.free;
  gs.DeleteRow (Gs.row);
  GSRowEnter(self,1,cancel,false);

end;

procedure TOF_RECONDUCTION.BInsertClick(Sender: TObject);
var cancel : boolean;
begin
  Ok_Creation := true;
  cancel := false;
  SetEnabled(True);
  TOBC := TOB.Create('BRECONDUCTION', TOBReconduction, -1);
  TOBC.AddChampSupValeur ('_NEW','-');
  TOBC.AddChampSupValeur ('_MODIFICATION','-',True);
  gs.Rowcount := gs.Rowcount+1;
  GS.row := TOBReconduction.detail.Count;
  GSRowEnter (self,GS.row,cancel,false);
end;

procedure TOF_RECONDUCTION.BACCEPTMODIFClick(Sender: TObject);
Var Code : String;
begin

    if Ok_Creation then
    begin;
	   Code := BRE_CODE.Text;
       if (Code = '') or (ControleCodeReconduction(Code) <> nil) then
	   begin
    	  BRE_CODE.Clear;
	      BRE_CODE.SetFocus;
    	  PGIBox(traduirememoire('Code reconduction invalide'), ecran.caption);
          exit;
       end;
    end;

	EnregistreModif;
	SetEnabled(false);

end;

procedure TOF_RECONDUCTION.ChampsExit(Sender: TObject);
var NewValue,NOM : string;
begin

	if Sender is THEdit Then
	begin
  	   NewValue := THEdit(Sender).Text;
	   Nom := THEDit(Sender).Name;
	end
    else if Sender is THValComboBox then
	begin
       NewValue := THValComboBox(Sender).Value;
	   Nom := THValComboBox(Sender).Name;
	end;

	if TOBC.GetValue (Nom) <> NewValue then TOBC.PutValue('_MODIFICATION','X');

end;

Function TOF_RECONDUCTION.ControleCodeReconduction(code : string) : TOB;
begin

   Result := TOBReconduction.FindFirst(['BRE_CODE'],[CODE], True);

end;

Function TOF_RECONDUCTION.ControleAffaire : boolean;
var QQ : TQuery;
begin
	QQ := OpenSQL('SELECT AFF_RECONDUCTION FROM AFFAIRE WHERE AFF_RECONDUCTION = "' + BRE_CODE.text + '"', True,-1,'',true);
	Result := QQ.eof;
    ferme (QQ);
end;

procedure TOF_RECONDUCTION.GSEnter(Sender: TObject);
var cancel : boolean;
begin
  cancel := false;
  GSRowEnter(self,1,cancel,false);
end;

procedure TOF_RECONDUCTION.GSRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin
  TOBC := TOBReconduction.detail[GS.row-1];
  TOBC.PutEcran (ecran);
end;

procedure TOF_RECONDUCTION.GSRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
var laTOB,TOBSAV : TOB;
begin
  LaTOB := TOBReconduction.detail[Ou-1];
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

function TOF_RECONDUCTION.ModifAutorise : boolean;
begin
	result := true;
end;

procedure TOF_RECONDUCTION.GSDblClick(Sender: TObject);
var IsModifAutorise : boolean;
begin
  Ok_Creation := False;
  IsModifAutorise := ModifAutorise;
  if ModifAutorise then SetEnabled (True);
end;

procedure TOF_RECONDUCTION.BANNULMODIFClick(Sender: TObject);
var Cancel : boolean;
begin
  Cancel := false;
  if ok_creation then
  begin
    TOBC.Free;
    gs.Rowcount := tobreconduction.Detail.count+1;
    GS.row := gs.RowCount-1;
    AffichelaLigne(GS.row);
    GSRowEnter (self,GS.row,cancel,false);
  end
  else
  begin
  	TOBC := TOBReconduction.Detail[GS.row-1];
	AffichelaLigne(GS.row);
	GSRowEnter (self,GS.row,cancel,false);
	TOBC.PutValue('_MODIFICATION','-');
  end;

  SetEnabled(false);
  Ok_Creation := false;

end;

procedure TOF_RECONDUCTION.AffichelaLigne(LaLigne: integer);
var TOBC : TOB;
begin
  TOBC := TOBReconduction.detail[LaLigne-1];
  TOBC.PutLigneGrid (GS,LaLigne,false,false,'BRE_CODE;BRE_LIBELLE');
end;

procedure TOF_RECONDUCTION.EnregistreModif;
begin
  TOBC.GetEcran (ecran);
  BEGINTRANS;
  TRY
    if not TOBC.InsertOrUpdateDB then V_PGI.IOError := OEUnknown;
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

procedure TOF_RECONDUCTION.MiseAjourTablette;
var TOBCC : TOB;
begin
	TOBCC := TOB.Create ('COMMUN',nil,-1);
  TRY
    TOBCC.PutValue('CO_TYPE','RCA');
    TOBCC.PutValue('CO_CODE',TOBC.GetValue('BRE_CODE'));
    TOBCC.PutValue('CO_LIBELLE',TOBC.GetValue('BRE_LIBELLE'));
	TOBCC.PutValue('CO_ABREGE','');
    TOBCC.PutValue('CO_LIBRE','---');
    if not TOBCC.InsertOrUpdateDB (false) Then V_PGI.IOError := oeUnknown;
  FINALLY
  TOBCC.free;
  END;
end;

procedure TOF_RECONDUCTION.SupprimeTablette;
Begin
  ExecuteSql('DELETE COMMUN WHERE CO_TYPE = "RCA" AND CO_CODE = "' + BRE_CODE.Text + '"');
End;

procedure TOF_RECONDUCTION.SupprimeEnreg(TOBS: TOB);
begin
  TOBS.DeleteDB (false);
end;      

Initialization
  registerclasses ( [ TOF_RECONDUCTION ] ) ;
end.

