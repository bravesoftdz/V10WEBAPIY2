{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 24/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTREPARTTVA ()
Mots clefs ... : TOF;BTREPARTTVA
*****************************************************************}
Unit BTREPARTTVA_TOF ;

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
     HTB97,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOB,
     EntGc,
     AglInit,
     UTOF,
     vierge,
     Graphics ;

Type

  TOF_BTREPARTTVA = Class (TOF)
      procedure OnNew                    ; override ;
      procedure OnDelete                 ; override ;
      procedure OnUpdate                 ; override ;
      procedure OnLoad                   ; override ;
      procedure OnArgument (S : String ) ; override ;
      procedure OnDisplay                ; override ;
      procedure OnClose                  ; override ;
      procedure OnCancel                 ; override ;
  	private
    	FromAnnul : boolean;
    	TOBRepart,TOBRepart_O : TOB;
      LTAXE1,LTAXE2,LTAXE3,LTAXE4,LTAXE5 : THLabel;
      TAXE1,TAXE2,TAXE3,TAXE4,TAXE5 : THValcomboBox;
      USABLE1,USABLE2,USABLE3,USABLE4,USABLE5 : TCheckBox;
      Bferme,Bdelete : TToolBarButton97;
      procedure GetComponents;
      procedure SetComponents;
  		procedure SetValues;
      procedure BfermeClick (Sender : TObject);
      procedure BDeleteClick (Sender : TObject);
      function ControleRepartition : boolean;
  		procedure EnregistreRepartition;
  end ;

procedure definiRepartitionTva (action : TactionFiche);

Implementation
uses Facture,FactTOB;

procedure definiRepartitionTva (action : TactionFiche);
begin
	AglLanceFiche ('BTP','BTREPARTTVA','','',ActionToString (Action));
end;

procedure TOF_BTREPARTTVA.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTREPARTTVA.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTREPARTTVA.OnUpdate ;
begin
  Inherited ;
  if ControleRepartition then
  begin
  	EnregistreRepartition
  end else
  begin
  	ecran.ModalResult := mrNone;
    exit;
  end;
  TheTOb := TOBrepart;
  TheTOB.PutValue ('VALIDATION','X');
end ;

procedure TOF_BTREPARTTVA.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTREPARTTVA.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBRepart := LaTOB;
  TOBRepart_O := TOB.Create ('LA REPARTITION TVA',nil,-1);
  TOBRepart_O.Dupliquer (TOBRepart,true,true);
  GetComponents;
  SetComponents;
  SetValues;
  FromAnnul := false;
end ;

procedure TOF_BTREPARTTVA.OnClose ;
begin
  Inherited ;
  TOBRepart_O.free;
end ;

procedure TOF_BTREPARTTVA.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTREPARTTVA.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTREPARTTVA.GetComponents;
begin
  LTAXE1 := THLabel (GetControl('LTAXE1'));
  LTAXE2 := THLabel (GetControl('LTAXE2'));
  LTAXE3 := THLabel (GetControl('LTAXE3'));
  LTAXE4 := THLabel (GetControl('LTAXE3'));
  LTAXE5 := THLabel (GetControl('LTAXE4'));
  TAXE1 := THValComboBox (GetControl('TAXE1'));
  TAXE2 := THValComboBox (GetControl('TAXE2'));
  TAXE3 := THValComboBox (GetControl('TAXE3'));
  TAXE4 := THValComboBox (GetControl('TAXE4'));
  TAXE5 := THValComboBox (GetControl('TAXE5'));
  USABLE1 := TCheckBox(GetControl('USABLE1'));
  USABLE2 := TCheckBox(GetControl('USABLE2'));
  USABLE3 := TCheckBox(GetControl('USABLE3'));
  USABLE4 := TCheckBox(GetControl('USABLE4'));
  USABLE5 := TCheckBox(GetControl('USABLE5'));
  BFerme := TToolBarButton97 (getControl('BFerme'));
  BDelete := TToolbarButton97(GetControl('Bdelete'));
end;

procedure TOF_BTREPARTTVA.SetComponents;
var Indice : integer;
		UneTaxe : TOB;
begin
	for Indice := 0 to VH_GC.TOBParamTaxe.detail.count -1 do
  begin
    Unetaxe  := VH_GC.TOBParamTaxe.detail[Indice];
    // non paramétré
    THLabel (GetControl('LTAXE'+inttostr(Indice+1))).Visible := false;
    THValComboBox (GetControl('TAXE'+inttostr(Indice+1))).Visible := false;
    TCheckBox (GetControl('USABLE'+inttostr(Indice+1))).Checked := false;
    THNumedit (GetControl('REPART'+inttostr(Indice+1))).visible := false;
    if UneTaxe.getValue('BPT_TYPETAXE') <> '' then
    begin
      if UneTaxe.getValue('BPT_TYPETAXE') = 'TVA' Then   
      begin
      	THValComboBox (GetControl('TAXE'+inttostr(Indice+1))).DataType := 'TTTVA';
      	TCheckBox (GetControl('USABLE'+inttostr(Indice+1))).Checked  := True;
      	THLabel (GetControl('LTAXE'+inttostr(Indice+1))).Caption := UneTaxe.getValue('BPT_LIBELLE');
      	THLabel (GetControl('LTAXE'+inttostr(Indice+1))).visible := true;
      	THValComboBox (GetControl('TAXE'+inttostr(Indice+1))).Visible := True;
    		THNumedit (GetControl('REPART'+inttostr(Indice+1))).visible := True;
    		THNumedit (GetControl('REPART'+inttostr(Indice+1))).Enabled := True;
      end else
      begin if UneTaxe.getValue('BPT_TYPETAXE') = 'TPF' then
(*
      	THValComboBox (GetControl('TAXE'+inttostr(Indice+1))).DataType := 'TTTPF';
      	TCheckBox (GetControl('USABLE'+inttostr(Indice+1))).Checked  := false;
      	THValComboBox (GetControl('TAXE'+inttostr(Indice+1))).Enabled := false;
      	THValComboBox (GetControl('TAXE'+inttostr(Indice+1))).Visible := True;
      	THLabel (GetControl('LTAXE'+inttostr(Indice+1))).Caption := UneTaxe.getValue('BPT_LIBELLE');
      	THLabel (GetControl('LTAXE'+inttostr(Indice+1))).visible := true;
    		THNumedit (GetControl('REPART'+inttostr(Indice+1))).visible := True;
    		THNumedit (GetControl('REPART'+inttostr(Indice+1))).Enabled := false;
    		THNumedit (GetControl('REPART'+inttostr(Indice+1))).color := clInactiveCaptionText;
*)
      end;
    end;
  end;
  Bferme.OnClick := BfermeClick;
(*
  if Tobrepart.detail.count > 0  Then
  begin
  	Bdelete.Visible := true;
    Bdelete.onclick := BdeleteClick;
  end;
*)
end;


procedure TOF_BTREPARTTVA.SetValues;
var indice : integer;
		TOBR : TOB;
    NumTva : string;
begin
	for indice := 0 to TOBrepart.detail.count -1 do
  begin
    TOBR := TOBrepart.detail[Indice];
    NumTva := Copy (TOBR.GetValue('BPM_CATEGORIETAXE'),3,1);
    THValComboBox (GetControl('TAXE'+NumTva)).Value := TOBR.GetValue('BPM_FAMILLETAXE');
    THNumedit (GetControl('REPART'+NumTva)).Value := TOBR.GetValue('BPM_MILLIEME');
  end;
end;

procedure TOF_BTREPARTTVA.BfermeClick(Sender: TObject);
begin
   TOBrepart.Dupliquer (TOBrepart_O,true,true); // annulation
end;

function TOF_BTREPARTTVA.ControleRepartition: boolean;
var Indice : integer;
		TaxeUsed,EnCours : string;
    RepartTot : double;
    messageError : string;
begin

	result := true;
	TaxeUsed := '';
  repartTot := 0;
	for Indice := 0 to VH_GC.TOBParamTaxe.detail.count -1 do
  begin
  	if TCheckBox (getControl('USABLE'+inttoStr(Indice+1))).Checked  then
    begin
    	EnCours := THValComboBox (GetControl('TAXE'+inttostr(Indice+1))).Value;
      if Encours <> '' then
      begin
        if Pos (Encours,TaxeUsed) > 0 then
        BEGIN
        	messageError := TraduireMemoire('La TVA ')+
          								THValComboBox (GetControl('TAXE'+inttostr(Indice+1))).Text+
                          Traduirememoire(' est déjà utilisée');
          PGIError (MessageError);
          THValComboBox (GetControl('TAXE'+inttostr(Indice+1))).SetFocus;
        	result := false;
          Exit;
        END;
        if (THNumEdit(getControl('REPART'+IntToStr(Indice+1))).Value = 0) and (not FromAnnul) then
        begin
        	messageError := TraduireMemoire('La répartition de la TVA ')+
          								inttostr(Indice+1)+
                          Traduirememoire(' n''est pas définie');
          PGIError (MessageError);
          THNumEdit (GetControl('REPART'+inttostr(Indice+1))).SetFocus;
        	result := false;
          Exit;
        end;
        TaxeUsed := TaxeUsed + EnCours+ ';';
        repartTot := RepartTot + THNumEdit (GetControl('REPART'+inttostr(Indice+1))).Value;
      end;
    end;
  end;

  if (RepartTot <> 1000) and (RepartTot <> 0) then
  begin
    messageError := TraduireMemoire('La répartition de la TVA n''est pas égale à 1000');
    PGIError (MessageError);
    result := false;
  end;
end;

procedure TOF_BTREPARTTVA.EnregistreRepartition;
var Indice : integer;
		EnCours : string;
    Repart : double;
    TOBR : TOB;
begin
	TOBrepart.clearDetail;
	for Indice := 0 to VH_GC.TOBParamTaxe.detail.count -1 do
  begin
  	if TCheckBox (getControl('USABLE'+inttoStr(Indice+1))).Checked  then
    begin
    	EnCours := THValComboBox (GetControl('TAXE'+inttostr(Indice+1))).Value;
      Repart :=  THNumEdit (GetControl('REPART'+inttostr(Indice+1))).Value;
      if (Encours <> '') and (Repart <> 0) then
      begin
        TOBR := TOB.Create ('BTPIECEMILIEME',TOBrepart,-1);
        TOBR.InitValeurs;
        TOBR.PutValue('BPM_CATEGORIETAXE','TX'+InttoStr(Indice+1));
        TOBR.PutValue('BPM_FAMILLETAXE',EnCours);
        TOBR.PutValue('BPM_MILLIEME',Repart);
      end;
    end;
  end;
end;

procedure TOF_BTREPARTTVA.BDeleteClick(Sender: TObject);
var Indice : integer;
		UneTaxe : TOB;
begin
	for Indice := 0 to VH_GC.TOBParamTaxe.detail.count -1 do
  begin
    Unetaxe  := VH_GC.TOBParamTaxe.detail[Indice];

    // non paramétré
    if UneTaxe.getValue('BPT_TYPETAXE') <> '' then
    begin
      if (UneTaxe.getValue('BPT_TYPETAXE') = 'TVA')  Then
      begin
      	if Indice > 0 then THValComboBox (GetControl('TAXE'+inttostr(Indice+1))).value := '';
    		THNumedit (GetControl('REPART'+inttostr(Indice+1))).value := 0;
      end else
      begin if UneTaxe.getValue('BPT_TYPETAXE') = 'TPF' then
      end;
    end;
  end;
  FromAnnul := true;
  TToolbarButton97 (GetControl('Bvalider')).Click;
end;

Initialization
  registerclasses ( [ TOF_BTREPARTTVA ] ) ;
end.

