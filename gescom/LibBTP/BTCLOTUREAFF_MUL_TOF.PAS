{***********UNITE*************************************************
Auteur  ...... : LS
Créé le ...... : 16/03/2009
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCLOTUREAFF_MUL ()
Mots clefs ... : TOF;BTCLOTUREAFF_MUL
*****************************************************************}
Unit BTCLOTUREAFF_MUL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,fe_main,
{$else}
     eMul,MainEagl,
{$ENDIF}
		 AglInit,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF ,
     HTB97,
     HDB,
     splash,
     utofAfBaseCodeAffaire,
     BTCLOTUREAFF_TOF;

Type
  TOF_BTCLOTUREAFF_MUL = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;
  private
  	Reactivation: boolean;
    BCherche    : TToolbarButton97;
    BOuvrir1		: TToolbarButton97;

    Fliste 			: THDbGrid;
    TOBChantiers: TOB;
    DateCloture : TDateTime;
    procedure MultiAcceptation(Sender: TOBJect);
    procedure FlisteDblClick(Sender: TObject);
    procedure ClotureChantierSolo(TheAffaire: string);
    procedure DeClotureChantierSolo(TheAffaire: string);
    procedure BVOIRClick(Sender: TObject);
  end ;

Implementation

uses uTOFComm;

procedure TOF_BTCLOTUREAFF_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTCLOTUREAFF_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTCLOTUREAFF_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTCLOTUREAFF_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTCLOTUREAFF_MUL.OnArgument (S : String ) ;
begin
  fMulDeTraitement := True;
  Inherited ;
  Reactivation := (Pos('REACTIVE',S)>0);
  if reactivation then BEGIN ecran.caption := 'Réactivation des affaires'; UpdateCaption (ecran); END;
// ----- FQ 12738 ----
	Fliste := THDbGrid (GetControl('FLISTE'));
  TToolBarButton97(GetControl('BVOIR')).onclick := BVOIRClick;
  //
  BCherche := TtoolBarButton97(GetControl('BCHERCHE'));
  //
  BOuvrir1 := TtoolBarButton97(GetControl('BOUVRIR1'));
  if not reactivation then
  begin
  	BOuvrir1.OnClick := MultiAcceptation;
  end else
  begin
  	fliste.MultiSelection := false;
  	Fliste.OnDblClick := FlisteDblClick;
  end;
  TOBChantiers := TOB.Create ('LES CHANTIERS A CLOTURER',nil,-1);
  TOBChantiers.AddChampSupValeur ('MODE','MULTI');
  TOBChantiers.AddChampSupValeur ('OK','-');
end ;

procedure TOF_BTCLOTUREAFF_MUL.OnClose ;
begin
  TOBChantiers.free;
  Inherited ;
end ;

procedure TOF_BTCLOTUREAFF_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTCLOTUREAFF_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTCLOTUREAFF_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2,
  Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin

Aff:=THEdit(GetControl('AFF_AFFAIRE'));

// MODIF LS
Aff0 := THEdit(GetControl('AFF_AFFAIRE0'));

// --
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
Tiers:=THEdit(GetControl('AFF_TIERS'));

// affaire de référence pour recherche
Aff_:=THEdit(GetControl('AFF_AFFAIREREF'));
Aff1_:=THEdit(GetControl('AFFAIREREF1'));
Aff2_:=THEdit(GetControl('AFFAIREREF2'));
Aff3_:=THEdit(GetControl('AFFAIREREF3'));
Aff4_:=THEdit(GetControl('AFFAIREREF4'));

end;

Procedure TOF_BTCLOTUREAFF_MUL.MultiAcceptation(Sender : TOBJect);
var TheAffaire : String;
    i,retour  : integer;
    QQ			 : TQuery;
{$IFDEF EAGLCLIENT}
    L 			 : THGrid;
{$ELSE}
    L 			 : THDBGrid;
{$ENDIF}
    uneTOB      : TOB;
    bConfirm    : boolean;
begin
  Inherited ;

  //controle si au moins un éléments sélectionné
	if (TFmul(ecran).FListe.NbSelected=0)and(not TFmul(ecran).FListe.AllSelected) then
  	 begin
	   PGIInfo('Aucun chantier sélectionné','');
  	 exit;
   	 end;

  //Demande de confirmation du traitement
  //uniquement en line
  //if PgiAsk ('Désirez-vous clôturer le(s) chantier(s) sélectionné(s) ?')<>MrYes then

  if PgiAsk ('Désirez-vous clôturer les affaires sélectionnées ?') <> MrYes then
     Begin
     SourisNormale;
     TFmul(ecran).BChercheClick(ecran);
  	 exit;
     end;

  bConfirm := false;

  DateCloture := Now;
  UneTOB := TOB.create ('UNE DEMANDE',nil,-1);
  UneTOB.AddChampSupValeur('CONFIRMATION','-');
  UneTOB.AddChampSupValeur('AFF_DATEFIN', DateCloture);
  UneTOB.AddChampSupValeur('MODECLOTURE','');
  TheTOB := UneTOB;

  TRY
    AGLLanceFiche('BTP','BTCONFIRMECLOAFF','','','');
    bConfirm := (UneTOB.GetValue('CONFIRMATION')='X');
    DateCloture := UneTob.GetDateTime('AFF_DATEFIN');
    if Datecloture = idate1900 then bconfirm := false;
    TheTOB := nil;
  Finally
  	UneTOB.free;
  end;
  if not bconfirm then
  begin
     SourisNormale;
     TFmul(ecran).BChercheClick(ecran);
  	 exit;
  end;
	L:= TFmul(ecran).FListe;

  SourisSablier;

  TRY
    if L.AllSelected then
    begin
      QQ:= TFmul(ecran).Q;
      QQ.First;
      while not QQ.EOF do
      Begin
        TheAffaire:= QQ.findfield('AFF_AFFAIRE').AsString;
        AddClotureChantierMul(TheAffaire,TOBChantiers,DateCloture);
      end;
      //
      QQ.next;
    end else
    begin
      for i:=0 to L.NbSelected-1 do
      begin
        L.GotoLeBookmark(i);
        TheAffaire := TFMul(TFmul(ecran)).Fliste.datasource.dataset.FindField('AFF_AFFAIRE').AsString;
        AddClotureChantierMul(TheAffaire,TOBChantiers,DateCloture);
      end;
    end;
    //FV1 : 11/09/2013 - FS#654 - BAGE : en cloture de chantier, ajouter la saisie de la date de clôture
    LanceClotureChantierMul(ecran,TOBChantiers,DateCloture);
    //
    if TOBChantiers.getValue('OK')='X' then
    begin
      retour :=  lancetraitementcloturechantierMul (ecran,TOBCHantiers, DateCloture);
      if retour > 0 then
      begin
        PgiInfo ('Les chantiers souhaités ont été clôturés');
      end else if (retour < 0) then
      begin
        PgiError ('Erreur durant le traitement de clôture');
      end;
    end;
	FINALLY
  	TOBChantiers.cleardetail;
    TOBChantiers.putvalue('OK','-');
    L.AllSelected:=False;
    SourisNormale;
    TFmul(ecran).BChercheClick(ecran);
  END;

End;

procedure TOF_BTCLOTUREAFF_MUL.FlisteDblClick (Sender : TObject);
var TheAffaire : string;
begin

  TheAffaire:= TFmul(ecran).Q.FindField('AFF_AFFAIRE').AsString;
  if not Reactivation then
  begin
//uniquement en line
//  if PgiAsk ('Désirez-vous clôturer le chantier ?')=MrYes then
    if PgiAsk ('Désirez-vous clôturer l''affaire ?')=MrYes then
    ClotureChantierSolo(TheAffaire);
  end else
  begin
//uniquement en line
//  if PgiAsk ('Désirez-vous réactiver ce chantier ?')=MrYes then
    if PgiAsk ('Désirez-vous réactiver cette affaire ?')=MrYes then
    DeClotureChantierSolo(TheAffaire);
  end;
  TFmul(ecran).BChercheClick(ecran);
end;

procedure TOF_BTCLOTUREAFF_MUL.DeClotureChantierSolo(TheAffaire: string);
var splash  : TFsplashScreen;
		OkTrait : integer;
    Ureason : integer;
begin
  	BEGINTRANS;
    TRY
  		LanceDeClotureSolo (TheAffaire);
      PgiInfo ('Chantier réactivé avec succès');
      COMMITTRANS;
    EXCEPT
      PgiInfo ('Erreur en cours de réactivation');
    	ROLLBACK;
    end;
end;

procedure TOF_BTCLOTUREAFF_MUL.ClotureChantierSolo(TheAffaire: string);
var splash  : TFsplashScreen;
		OkTrait : integer;
    Ureason : integer;
begin
  //
  splash := TFsplashScreen.Create (fliste.Parent);
  splash.Label1.Caption := 'Controle du chantier '+ TheAffaire+' en cours...';
  splash.Show;
  splash.Refresh;
  OkTrait :=  ControleChantierSolo(TheAffaire,Ureason,DateCloture);
  splash.free;
  if OkTRait>=1 then
  begin
  	if AutoriseClotureChantier (TheAffaire,Ureason, DateCloture) then
    begin
    	OkTrait := 0;
    end;
  end;
  if OkTrait=0 then
  begin
  	BEGINTRANS;
    TRY
  		LanceTraitementClotureSolo (TheAffaire, DateCloture);
      PgiInfo ('Chantier clôturé avec succès');
      COMMITTRANS;
    EXCEPT
      PgiInfo ('Erreur en cours de clôture');
    	ROLLBACK;
    end;
  end;
end;


procedure TOF_BTCLOTUREAFF_MUL.BVOIRClick(Sender: TObject);
var Affaire				: String;
begin
	 if FListe.datasource.DataSet.RecordCount = 0  then exit;

   Affaire:=Fliste.datasource.dataset.FindField('AFF_AFFAIRE').AsString;

	 AGLLanceFiche ('BTP','BTAFFAIRE','',Affaire,'ACTION=CONSULTATION');
end;

Initialization
  registerclasses ( [ TOF_BTCLOTUREAFF_MUL ] ) ;
end.
