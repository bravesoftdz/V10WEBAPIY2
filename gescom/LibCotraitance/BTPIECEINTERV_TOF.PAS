{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/10/2011
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPIECEINTERV ()
Mots clefs ... : TOF;BTPIECEINTERV
*****************************************************************}
Unit BTPIECEINTERV_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,MainEagl,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     HTB97,
     Vierge,
     UCotraitance,
     BTPUtil ;

Type
  TOF_BTPIECEINTERV = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    BPI_NUMEROCONTACT : Thedit;
    BPI_TIERSFOU      : Thedit;
    BPI_NUMERORIB     : THEdit;
    BPI_TYPEPAIE      : THValComboBox;
    BPI_FAMILLETAXE   : THvalComboBox;
    BPI_AUTOLIQUID    : THCheckBox; 
    //
    FromlaTOB : Boolean;
    TOBssTrait,TOBDESSOUSTRAIT : TOB;
    TheAction : string;
    ErrorMsg : string;
    Naturepieceg,Souche,Fournisseur,Affaire : string;
    Numero,Indiceg : Integer;
    procedure DecodeArgs(S: string);
    procedure InitEnreg;
    procedure LoadLeSousTraitant;
    procedure Charge_InfoBQE(TOBssTrait,TOBBQE: TOB);
    procedure SetInfoBqe;
    procedure ChangeTzTiersSaisie(TypeCompte: string);
    procedure TIERSFOUElipsisClick(Sender: TObject);
    procedure TIERSFOUEnter(Sender: Tobject);
    procedure TIERSFOUEXit(Sender: Tobject);
    procedure ContactClick(Sender: TObject);
    procedure ContactExit(Sender: tobject);
    procedure RIBElipsisClick(Sender: TObject);
    procedure CodeBQExit(Sender: TObject);
    procedure CodeBQOnChange(Sender: TObject);
    procedure FamilleTaxeChange (Sender : TObject);
    function ControleExist(TOBL: TOB; Fournisseur,TypeIntervenant: string): integer;
    procedure SetNewNumOrdre;
    procedure BdeleteClick (Sender : Tobject);
    function OnBeforeUpdate(var TypeErreur: string): Boolean;
    procedure TypePaieExit (Sender : Tobject);
    function ControleRib: Boolean;
    procedure ChargeInfoBqe(TOBssTrait: TOB);
    function ExistsSoustraitInTOB(Fournisseur : string; TOBDESSOUSTRAIT : TOB) : Boolean;
    procedure BvaliderClick (Sender : Tobject);
    procedure ControleEtValide;

  end ;

Implementation
uses TiersUtil,FactTiers,ENtGC,ParamSoc;

procedure TOF_BTPIECEINTERV.OnNew ;
begin
  Inherited ;
  THCheckBox(getControl('BPI_AUTOLIQUID')).Checked := true;
  THvalComboBox(getControl('BPI_FAMILLETAXE')).Value := GetParamSocSecur('SO_BTTAXESOUSTRAIT','TN'); 
end ;

function TOF_BTPIECEINTERV.OnBeforeUpdate(var TypeErreur : string) : Boolean ;
var ferr : Integer;
begin

  Result := True;

  if (BPI_NUMEROCONTACT.Text = '') OR (BPI_NUMEROCONTACT.Text = '0') then
  begin
    BPI_NUMEROCONTACT.SetFocus;
    ErrorMsg := 'Attention le contact n''est pas renseigné.' + CHR(10) + 'Confirmez-vous la validation';
    TypeErreur := 'N';
    Result := False;
    Exit;
  end;

 	if getControltext('BPI_TIERSFOU') = '' then
  begin
  	PgiInfo ('Veuillez renseigner le code du sous traitant.');
    BPI_TIERSFOU.SetFocus;
    TypeErreur := 'B';
    Result := False;
    Exit;
  end;

  if TheAction = 'CREATION' then
  begin
    ferr :=  ControleExist (TOBssTrait,BPI_TIERSFOU.text,GetControltext('BPI_TYPEINTERV'));
    if ferr = 1 then
    begin
      PgiInfo ('impossible : Cet intervenant est déjà défini');
      BPI_TIERSFOU.text := TOBssTrait.getvalue('BPI_TIERSFOU');
      BPI_TIERSFOU.SetFocus;
      TypeErreur := 'B';
    	Result := False;
      Exit;
    end;
    if Ferr = 2 then
    begin
      PgiInfo ('impossible : Cette société est déjà défini comme co-traitante');
      BPI_TIERSFOU.SetFocus;
      TypeErreur := 'B';
    	Result := False;
      Exit;
    end;
  end;
  //

 	if getControltext('BPI_TYPEPAIE') = '' then
  begin
  	PgiInfo ('Veuillez renseigner le type de règlement du sous traitant ');
    BPI_TYPEPAIE.SetFocus;
    TypeErreur := 'B';
    Result := False;
    Exit;
  end;

  if BPI_TYPEPAIE.Value = '001' then
  begin
    if not ControleRib then
    begin
      PgiInfo ('Cette banque n''existe pas. Merci d''en sélectionner une autre');
      BPI_NUMERORIB.Text := '0';
      BPI_NUMERORIB.SetFocus;
      TypeErreur := 'B';
    	Result := False;
      Exit;
    end;
  end;
end;

procedure TOF_BTPIECEINTERV.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECEINTERV.OnUpdate ;
begin
  Inherited ;
  //
  if not FromlaTOB then
  begin
  	ControleEtValide;
  end;
end ;

procedure TOF_BTPIECEINTERV.OnLoad ;
begin
  Inherited ;
  SetControlText('TLIB_NUMERORIB', TOBssTrait.GetString('TLIB_NUMERORIB'));
end ;

procedure TOF_BTPIECEINTERV.OnArgument (S : String ) ;
begin
  Inherited ;
  //
  BPI_TIERSFOU      := Thedit(getControl('BPI_TIERSFOU'));
  BPI_NUMEROCONTACT := Thedit(getControl('BPI_NUMEROCONTACT'));
  BPI_NUMERORIB     := THEdit(GetControl('BPI_NUMERORIB'));
  BPI_TYPEPAIE      := THvalComboBox(getControl('BPI_TYPEPAIE'));
  BPI_FAMILLETAXE   := THvalComboBox(getControl('BPI_FAMILLETAXE'));
  BPI_AUTOLIQUID    := THCheckBox(getControl('BPI_AUTOLIQUID'));
  //
  if LaTOB <> nil then
  begin
    TOBssTrait := LaTOB;
    TOBDESSOUSTRAIT := TOB(TOBssTrait.data);
		ChargeInfoBqe (TOBssTrait);
    FromlaTOB := True;
  end else
  begin
    FromlaTOB := false;
    TOBssTrait := TOB.Create('PIECEINTERV',nil,-1);
		AddChampsSupTraitCreat (TOBSSTRAIT);
		AddChampsSupTrait (TOBSSTRAIT);
  end;
  DecodeArgs (S);
  if TheAction = 'CREATION' then
  begin
  	InitEnreg;
    TOBssTrait.PutEcran(Ecran);
    SetNewNumOrdre;
  end else if TheAction = 'MODIFICATION' then
  begin
    if not FromlaTOB then LoadLeSousTraitant;
    TOBssTrait.PutEcran(Ecran);
    BPI_TIERSFOU.Enabled := false;
  end;
  SetInfoBqe;
  SetControltext('BPI_DATECREATION',DateTimeToStr(TOBssTrait.getValue('BPI_DATECREATION')));
  SetControltext('TBPI_CREATEUR',Rechdom('TTUTILISATEUR',TOBssTrait.getValue('BPI_CREATEUR'),false));
  SetControltext('BPI_DATEMODIF',DateTimeToStr(TOBssTrait.getValue('BPI_DATEMODIF')));
  SetControltext('TBPI_UTILISATEUR',Rechdom('TTUTILISATEUR',TOBssTrait.getValue('BPI_UTILISATEUR'),false));
  //
  BPI_TIERSFOU.OnEnter              := TIERSFOUEnter;
  BPI_TIERSFOU.OnExit               := TIERSFOUEXit;
  BPI_TIERSFOU.OnElipsisClick       := TIERSFOUElipsisClick;
  //
  BPI_NUMEROCONTACT.OnElipsisClick  := ContactClick;
  BPI_NUMEROCONTACT.OnExit          := ContactExit;
  //
  BPI_NUMERORIB.OnElipsisClick      := RIBElipsisClick;
  BPI_NUMERORIB.OnChange            := CodeBQOnChange;
  BPI_NUMERORIB.OnExit              := CodeBQExit;
  BPI_FAMILLETAXE.OnChange          := FamilleTaxeChange;
  TToolbarButton97 (GetControl('Bdelete')).OnClick := BdeleteClick;
  BPI_TYPEPAIE.OnChange := TypePaieExit;
  if FromlaTOB then TToolbarButton97 (GetControl('BValider')).OnClick := BvaliderClick;
  //
end ;

procedure TOF_BTPIECEINTERV.OnClose ;
begin
	if not FromlaTOB then TOBssTrait.Free;
  Inherited ;
end ;

procedure TOF_BTPIECEINTERV.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECEINTERV.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECEINTERV.DecodeArgs(S: string);
var LesArgs,UnArg : string;
		UnCode,Unevaleur : string;
    IposEgal : integer;
begin
	lesArgs := S;
  Repeat
  	UnArg := READTOKENST(LesArgs);
    if UnArg <> '' then
    begin
      IposEgal := Pos('=',UnArg);
      if IposEgal > 0 then
      begin
        Uncode := copy(unArg,1,IPosEgal-1);
        UneValeur := copy(unArg,IPosEgal+1,255);
        if Uncode = 'BPI_NATUREPIECEG' then
        begin
          Naturepieceg := Unevaleur;
        end else if Uncode = 'BPI_SOUCHE' then
        begin
          Souche:= Unevaleur;
        end else if Uncode = 'BPI_NUMERO' then
        begin
          Numero:= StrToInt(Unevaleur);
        end else if Uncode = 'BPI_INDICEG' then
        begin
          Indiceg := StrToInt(Unevaleur);
        end else if Uncode = 'BPI_TIERSFOU' then
        begin
           Fournisseur := Unevaleur;
        end else if Uncode = 'ACTION' then
        begin
           TheAction := Unevaleur;
        end else if Uncode = 'AFFAIRE' then
        begin
           Affaire := Unevaleur;
        end;
      end;
    end;
  Until UnArg='';
end;

procedure TOF_BTPIECEINTERV.InitEnreg;
begin
  TOBssTrait.InitValeurs(false);
  TOBssTrait.putValue('BPI_NATUREPIECEG',Naturepieceg);
  TOBssTrait.putValue('BPI_SOUCHE',Souche);
  TOBssTrait.putValue('BPI_NUMERO',Numero );
  TOBssTrait.putValue('BPI_INDICEG',Indiceg);
  TOBssTrait.putValue('BPI_TYPEINTERV','Y00');
  TOBssTrait.putValue('BPI_TYPEPAIE','002');
  TOBssTrait.putValue('BPI_DATECONTRAT',V_PGI.DateEntree);
  TOBssTrait.SetBoolean('BPI_AUTOLIQUID',true);
  TOBssTrait.SetString('BPI_FAMILLETAXE',GetParamSocSecur('SO_BTTAXESOUSTRAIT','TN'));
end;

Procedure TOF_BTPIECEINTERV.Charge_InfoBQE(TOBssTrait,TOBBQE : TOB);
begin
  TOBssTrait.PutValue('ETABBQ', TOBBQE.GetString('R_ETABBQ'));
  TOBssTrait.PutValue('GUICHET',TOBBQE.GetString('R_GUICHET'));
  TOBssTrait.PutValue('NUMEROCOMPTE', TOBBQE.GetString('R_NUMEROCOMPTE'));
  TOBssTrait.PutValue('CLERIB', TOBBQE.GetString('R_CLERIB'));
  TOBssTrait.PutValue('IBAN', TOBBQE.GetString('R_CODEIBAN'));
end;

procedure TOF_BTPIECEINTERV.LoadLeSousTraitant;
var QQ : TQuery;
		Sql : string;
begin
  TOBssTrait.InitValeurs(false);
  Sql := 'SELECT *,'+
  			 '(SELECT T_LIBELLE FROM TIERS WHERE T_TIERS=BPI_TIERSFOU AND T_NATUREAUXI="FOU") AS TLIBTIERSFOU,'+
         '(SELECT R_DOMICILIATION FROM RIB WHERE R_AUXILIAIRE=BPI_TIERSFOU AND R_NUMERORIB=BPI_NUMERORIB) AS TLIB_NUMERORIB,' +
  			 '(SELECT C_NOM FROM CONTACT WHERE C_TYPECONTACT="T" AND C_NATUREAUXI="FOU" AND '+
         'C_AUXILIAIRE=(SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS=BPI_TIERSFOU AND T_NATUREAUXI="FOU") '+
         'AND C_NUMEROCONTACT=BPI_NUMEROCONTACT)  AS TLIBCONTACT '+
  			 'FROM PIECEINTERV WHERE BPI_NATUREPIECEG="'+naturePieceg+'" AND '+
  			 'BPI_SOUCHE="'+Souche+'" AND BPI_NUMERO='+IntToStr(Numero)+' AND BPI_INDICEG='+IntTOStr(Indiceg)+' AND '+
         'BPI_TIERSFOU="'+Fournisseur+'"';
  QQ := OpenSQL(Sql,True,1,'',true);
  TOBssTrait.SelectDB('',QQ);
  ferme (QQ);
  //
  ChargeInfoBqe (TOBssTrait);
end;

procedure TOF_BTPIECEINTERV.SetInfoBqe;
Begin

  SetControlText('TLIB_NUMERORIB', TOBssTrait.GetString('TLIB_NUMERORIB'));

  //chargement de la reférence bancaire
  if GetControlText('BPI_NUMERORIB') <> '0' then
    if (TOBssTrait.GetValue('ETABBQ')='') and (TOBssTrait.getvalue('GUICHET')='') and
    	 (TOBssTrait.GetValue('NUMEROCOMPTE')='') and (TOBssTrait.GetValue('CLERIB')='') then
      SetControltext('TLIBREFBANQUE', TOBssTrait.GetValue('IBAN'))
    else
      SetControltext('TLIBREFBANQUE', TOBssTrait.GetValue('ETABBQ') + ' - ' + TOBssTrait.getvalue('GUICHET') +
      							 ' - ' + TOBssTrait.GetValue('NUMEROCOMPTE') + '/' + TOBssTrait.GetValue('CLERIB'))
  else
    SetControltext('TLIBREFBANQUE', '');

end;

procedure TOF_BTPIECEINTERV.ChangeTzTiersSaisie(TypeCompte : string);
var St: string;
  ii: integer;
begin
  St := 'GCTIERSSAISIE';
  ii := TTToNum(St);
  if ii > 0 then
  begin
    V_PGI.DECombos[ii].Libelle := 'Fournisseurs';
    V_PGI.DECombos[ii].Where := ' AND (T_NATUREAUXI="FOU")';
  end;
end;


procedure TOF_BTPIECEINTERV.TIERSFOUEnter(Sender: Tobject);
begin
	ChangeTzTiersSaisie ('FOU');
end;

procedure TOF_BTPIECEINTERV.TIERSFOUEXit(Sender: Tobject);
var LibTiers : string;
		ferr     : integer;
    StSQL    : string;
begin

  if GetTiersFerme('FOU',BPI_TIERSFOU.text,LibTiers) then
  begin
  	PgiInfo ('Ce Fournisseur est fermé. Merci d''en sélectionner un autre');
    SetControlText('BPI_TIERSOU',TOBssTrait.GetValue('BPI_TIERSFOU'));
    BPI_TIERSFOU.SetFocus;
    exit;
  end;

	if not GetLibTiers('FOU',BPI_TIERSFOU.text,LibTiers) then
  begin
  	PgiInfo ('Ce Fournisseur n''existe pas. Merci d''en sélectionner un autre');
    SetControlText('BPI_TIERSOU',TOBssTrait.GetValue('BPI_TIERSFOU'));
    BPI_TIERSFOU.SetFocus;
    exit;
  end else
  begin
  	ferr :=  ControleExist (TOBssTrait,BPI_TIERSFOU.text,'');
    if ferr = 1 then
    begin
    	PgiInfo ('impossible : Cet intervenant est déjà défini comme sous traitant');
      BPI_TIERSFOU.text := TOBssTrait.getvalue('BPI_TIERSFOU');
      BPI_TIERSFOU.SetFocus;
      exit;
    end else if ferr = 2 then
    begin
    	PgiInfo ('impossible : Cet intervenant est déjà défini comme cotraitant');
      BPI_TIERSFOU.text := TOBssTrait.getvalue('BPI_TIERSFOU');
      BPI_TIERSFOU.SetFocus;
      exit;
    end;
    SetControlText('TLIBTIERSFOU',LibTiers);
    TOBssTrait.putValue('TLIBTIERSFOU',LibTiers);
    //recherche des informations bancaires
    StSQL :='SELECT * FROM RIB WHERE R_AUXILIAIRE="' + BPI_TIERSFOU.text + '" AND R_PRINCIPAL="X"';
    //recupération du code banque par défaut....

  end;
end;

procedure TOF_BTPIECEINTERV.TIERSFOUElipsisClick(Sender: TObject);
begin
  if GetTiers(BPI_TIERSFOU) then
  begin
  	TIERSFOUEXit (self);
  end;
end;

procedure TOF_BTPIECEINTERV.ContactClick(Sender: TObject);
var retour,CodeAuxi,CodeTiers,Natureauxi,Action : string;
begin
  CodeTiers := BPI_TIERSFOU.Text;
  CodeAuxi  :=TiersAuxiliaire(CodeTiers,false,'FOU');
  NATUREAUXI:='FOU';
  Action    := 'ACTION=MODIFICATION';
  //
  retour := AglLanceFIche('YY','YYCONTACT','T;'+CodeAuxi,'',Action+ ';TYPE=T;'+'TYPE2='+Natureauxi+';TITRE=Sélection du destinataire;TIERS='+CodeTiers+';ALLCONTACT' );
  if retour <> '' then
  begin
  	THEdit(GetControl('BPI_NUMEROCONTACT')).Text := ReadTokenSt(Retour);
  end;
end;

procedure TOF_BTPIECEINTERV.ContactExit(Sender: tobject);
var libcontact,CodeTiers : string;
begin


  if StrToInt(BPI_NUMEROCONTACT.text) = 0 then Exit;

    // dans le cas de mandataire
  CodeTiers := BPI_TIERSFOU.Text;


	if not GetContact(CodeTiers ,StrToInt(BPI_NUMEROCONTACT.text),Libcontact) then
  begin
  	PgiInfo ('Ce contact n''existe pas. Merci d''en sélectionner un autre');
    SetControlText('BPI_NUMEROCONTACT',TOBssTrait.GetValue('BPI_NUMEROCONTACT'));
    BPI_NUMEROCONTACT.SetFocus;
    exit;
  end else
  begin
    SetControlText('TLIBCONTACT',LibContact);
    TOBssTrait.putValue('TLIBCONTACT',LibContact);
  end;

end;

procedure TOF_BTPIECEINTERV.RIBElipsisClick(Sender: TObject);
var CodeTiers,sNumeroRib,NatureAuxi,CodeAuxi : string;
begin

  CodeTiers := BPI_TIERSFOU.Text;
  NatureAuxi:='FOU';
  CodeAuxi := TiersAuxiliaire (CodeTIers,false,'FOU');

  //appel de la fiche RIB pour récupération code banque...
  sNumeroRIB := AGLLanceFiche('YY','YYRIB',CodeAuxi,'','NumAux='+ TiersAuxiliaire (CodeTIers,false,NatureAuxi));
  if sNumeroRib <> '' then
  begin
    BPI_NUMERORIB.Text := sNumeroRib;
    CodeBQExit(Self);
  end;

end;

procedure TOF_BTPIECEINTERV.CodeBQOnChange(Sender: TObject);
begin

  if BPI_NUMERORIB.Text <> '' then CodeBQExit (self);

end;


procedure TOF_BTPIECEINTERV.CodeBQExit(Sender: TObject);
Var RefBancaire : string;
    CodeTiers,CodeAuxi   : string;
    NumeroRIB   : Integer;
    TOBBQE      : TOb;
begin

  NumeroRIB := StrToInt(BPI_NUMERORIB.Text);
  if NumeroRIB = 0 then Exit;
    // dans le cas de mandataire
  CodeTiers := BPI_TIERSFOU.Text;
  CodeAuxi := TiersAuxiliaire (CodeTIers,false,'FOU');

  //recherche des information bancaire
	if not GetCodeBQ(CodeAuxi, NumeroRIB, RefBancaire) then
  begin
  	PgiInfo ('Cette banque n''existe pas. Merci d''en sélectionner une autre');
    BPI_NUMERORIB.Text := TOBssTrait.GetValue('BPI_NUMERORIB');
    exit;
  end else
  begin
    //Chargement suite info bancaire...
    TOBBQE := Tob.Create('RIB', nil, -1);
    LectRib(TOBBQE, CodeAuxi, NumeroRIB);
    TOBssTrait.putValue('TLIB_NUMERORIB',RefBancaire);
    if Assigned(TOBBQE) then
    begin
      TOBssTrait.putValue('ETABBQ', TOBBQE.GetString('R_ETABBQ'));
      TOBssTrait.putValue('GUICHET',TOBBQE.GetString('R_GUICHET'));
      TOBssTrait.putValue('NUMEROCOMPTE', TOBBQE.GetString('R_NUMEROCOMPTE'));
      TOBssTrait.putValue('CLERIB', TOBBQE.GetString('R_CLERIB'));
      TOBssTrait.putValue('IBAN', TOBBQE.GetString('R_CODEIBAN'));
    end
    else
    begin
      TOBssTrait.putValue('ETABBQ', '');
      TOBssTrait.putValue('GUICHET', '');
      TOBssTrait.putValue('NUMEROCOMPTE', '');
      TOBssTrait.putValue('CLERIB', '');
      TOBssTrait.putValue('IBAN', '');
    end;
    FreeandNil(TOBBQE);
    SetInfoBqe;
  end;

end;


function TOF_BTPIECEINTERV.ControleExist(TOBL: TOB;Fournisseur : string; TypeIntervenant : string): integer;
begin
	result := 0;
  if not FromlaTOB then
  begin
    if ExistsSoustrait(Fournisseur,Naturepieceg ,Souche,Numero,IndiceG) then
    begin
      result:= 1;
    end;
  end else
  begin
    if ExistsSoustraitInTOB(Fournisseur,TOBDESSOUSTRAIT) then
    begin
      result:= 1;
    end;
  end;
  if ExistsCotrait(Fournisseur,Affaire) then
  begin
    result:= 2;
  end;
end;

procedure TOF_BTPIECEINTERV.SetNewNumOrdre;
var QQ: TQuery;
		Num : integer;
    SQl : string;
    indice , Max : Integer;
begin
  if not FromlaTOB then
  begin
    Sql := 'SELECT MAX(BPI_ORDRE) AS LASTNUM '+
           'FROM PIECEINTERV WHERE BPI_NATUREPIECEG="'+naturePieceg+'" AND '+
           'BPI_SOUCHE="'+Souche+'" AND BPI_NUMERO='+IntToStr(Numero)+' AND BPI_INDICEG='+IntTOStr(Indiceg);
    QQ := OpenSQL(SQL,true,1,'',true);
    if not QQ.Eof then
    begin
      Num:= QQ.findField('LASTNUM').AsInteger;
      TOBssTrait.PutValue('BPI_ORDRE', Num+1 );
    end else
    begin
      TOBssTrait.PutValue('BPI_ORDRE',1);
    end;
    ferme (QQ);
  end else
  begin
    Max := 0;
		for indice := 0 to TOBDESSOUSTRAIT.Detail.count -1 do
    begin
      if TOBDESSOUSTRAIT.Detail[Indice].GetValue('BPI_ORDRE') > max then
				Max := TOBDESSOUSTRAIT.Detail[Indice].GetValue('BPI_ORDRE');
    end;
    TOBssTrait.PutValue('BPI_ORDRE', Max+1 );
  end;
end;

procedure TOF_BTPIECEINTERV.BdeleteClick(Sender: Tobject);
begin
   if PGIAsk('Etes-vous sur de vouloir supprimer ce sous-traitant ?') = mryes then
   begin
     TOBssTrait.DeleteDB(false);
     Ecran.Close;
   end;
end;

procedure TOF_BTPIECEINTERV.TypePaieExit(Sender: Tobject);
begin
	if BPI_TYPEPAIE.Value = '002' then
  begin
  	// paiement par nos soins ...donc pas besoin d'avoir le RIB du fournisseur ..se fera via la facture fournisseur
    SetControlVisible('TBPI_CODEBQ',false);
    SetControlVisible('BPI_NUMERORIB',false);
    SetControlVisible('TLIB_NUMERORIB',false);
    SetControlVisible('TLIBREFBANQUE',false);
    SetControlText('BPI_NUMERORIB','0');
  end else
  begin
    SetControlVisible('TBPI_CODEBQ',true);
    SetControlVisible('BPI_NUMERORIB',true);
    SetControlVisible('TLIB_NUMERORIB',true);
    SetControlVisible('TLIBREFBANQUE',true);
  end;
end;

function TOF_BTPIECEINTERV.ControleRib : Boolean;
var numerorib : Integer;
		CodeTiers,Refbancaire,CodeAuxi : string;
begin
  result := True;
  NumeroRIB := StrToInt(BPI_NUMERORIB.Text);
    // dans le cas de mandataire
  CodeTiers := BPI_TIERSFOU.Text;
  CodeAuxi := TiersAuxiliaire (CodeTiers,false,'FOU');


  //recherche des information bancaire
	if not GetCodeBQ(CodeAuxi, NumeroRIB, RefBancaire) then
  begin
    result := false;
  end;

end;

procedure TOF_BTPIECEINTERV.ChargeInfoBqe (TOBssTrait : TOB);
var TOBBQe : TOB;
		CodeAuxi : string;
begin
  TOBBQE := Tob.Create('RIB', nil, -1);
  CodeAuxi := TiersAuxiliaire (TOBssTrait.GetValue('BPI_TIERSFOU'),false,'FOU');
  LectRib(TOBBQE, CodeAuxi, TOBssTrait.GetValue('BPI_NUMERORIB'));
  Charge_InfoBqe(TOBssTrait,TOBBQE);
  freeAndNil(TOBBqe);
  SetInfoBqe ;
end;

function TOF_BTPIECEINTERV.ExistsSoustraitInTOB(Fournisseur: string;TOBDESSOUSTRAIT: TOB): Boolean;
begin
  result := (TOBDESSOUSTRAIT.findFirst(['BPI_TIERSFOU'],[Fournisseur],True)<>nil);
end;

procedure TOF_BTPIECEINTERV.BvaliderClick(Sender: Tobject);
begin
  ControleEtValide;
end;

procedure TOF_BTPIECEINTERV.ControleEtValide;
var TypeError : string;
begin
  if not OnBeforeUpdate(TypeError) then
  begin
    if TypeError = 'B' then
    begin
    	TFVierge(ecran).ModalResult := 0;
      Exit;
    end else if TypeError = 'N' then
    begin
      if PGIAsk(ErrorMsg, 'Erreur validation') = Mrno then
      begin
        TFVierge(ecran).ModalResult := 0;
      	Exit;
      end;
    end
    else
    begin
    	TFVierge(ecran).ModalResult := 0;
      PGIInfo('Erreur inconnue');
      Exit;
    end;
  end;

  TOBssTrait.GetEcran(ecran);
  TOBssTrait.SetAllModifie(true);
  if not FromlaTOB then
  begin
    if TheAction = 'CREATION' then
      TOBssTrait.InsertDB(nil,false)
    else
      TOBssTrait.UpdateDB(false);
  end;
end;

procedure TOF_BTPIECEINTERV.FamilleTaxeChange(Sender: TObject);
begin
//  if BPI_FAMILLETAXE.Value = VH_GC.AutoLiquiTVAST then BPI_AUTOLIQUID.Checked := true;
end;

Initialization
  registerclasses ( [ TOF_BTPIECEINTERV ] ) ;
end.
