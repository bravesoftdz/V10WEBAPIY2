{-------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
                17/09/01    RR     Création de l'unité
1.00.001.001    03/08/03    JP     Migration eAGL
1.00.001.001    02/09/03    JP     Gestion des agios/intérêts déduits
1.05.001.002    09/03/04    JP     Mise en place du Type de transaction dans les catégories de transaction
1.2X.000.000    05/04/04    JP     Correction du Bug sur les agios précomptés FQ 10034
6.xx.xxx.xxx    19/07/04    JP     Mise en place de commissions (cf routine "Validation")
6.30.001.002    09/03/05    JP     Suppression de la procédure Validation qui ne peut être exécutée car
                                   depuis la version 1, on ne peut plus valider depuis la fiche : régulièrement
                                   je maintenais le source inutilement cf. ci-dessus avec les commissions
7.05.001.001    18/09/06    JP     Gestion du NoDossier et de la société en Multi-sociétés
--------------------------------------------------------------------------------------}
unit TomCourtsTermes;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  {$IFDEF EAGLCLIENT}
  eFiche, MaineAGL, StdCtrls,
  {$ELSE}
  Fiche, FE_Main, DB, DBCtrls, HDB,
  {$ENDIF}
  Controls, Classes , Hctrls, ExtCtrls, Buttons, HEnt1, UTOB,
  UTOM, SysUtils;


procedure TRLanceFiche_CourtsTermes(Dom, Fiche, Range, Lequel, Arguments : string);

type
  TOM_CourtsTermes = class(TOM)
    procedure OnArgument(Arguments : string); override;
    procedure OnNewRecord                   ; override;
    procedure OnAfterUpdateRecord           ; override;
    procedure OnUpdateRecord        	    ; override;
    procedure OnLoadRecord                  ; override;
    procedure OnChangeField(F : TField)     ; override;
  private
    BCalc            	: TBitBtn ;
    BCalcTaux        	: TBitBtn ;
    BCalcMntAgios     	: TBitBtn ;
    lDev             	: THLabel ;
    lCodeTaux        	: THLabel ;
    lPlus            	: THLabel ;
    lMultiple        	: THLabel ;
    {$IFDEF EAGLCLIENT}
    dbCategorie      	: THEdit ;
    dbTransaction    	: THEdit ;
    cbContrePartie	: THEdit ;
    cbCompte		: THEdit ;
    dbValTaux        	: THEdit ;
    dbCodeTaux       	: THEdit ;
    dbMajoration    	: THEdit ;
    dbMultiple       	: THEdit ;
    dbMontant        	: THEdit ;
    dbMontantCtrVal	: THEdit ;
    dbMntMEP         	: THEdit ;
    dbMntAgios		: THEdit ;
    dbMntTombe		: THEdit ;
    dbDevMontant     	: THEdit ;
    dbDevCtrValeur	: THEdit ;
    cbTauxPrecompte     : TCheckBox;
    cbValBO          	: TCheckBox ;
    cbTauxFixe       	: TCheckBox ;
    cbAgiosDeduit    	: TCheckBox ;
    cbAgiosPrecompte    : TCheckBox ;
    dbTauxResultat   	: THEdit ;
    dDateDepart      	: THEdit ;
    dDateFin         	: THEdit ;
    dDateOpe         	: THEdit ;
    cbBaseCalcul	: THValComboBox ;
    spNbJBanque		: THSpinEdit ;
    {$ELSE}
    dbCategorie      	: THDBEdit ;
    dbTransaction    	: THDBEdit ;
    cbContrePartie	: THDBEdit ;
    cbCompte		: THDBEdit ;
    dbValTaux        	: THDBEdit ;
    dbCodeTaux       	: THDBEdit ;
    dbMajoration    	: THDBEdit ;
    dbMultiple       	: THDBEdit ;
    dbMontant        	: THDBEdit ;
    dbMontantCtrVal	: THDBEdit ;
    dbMntMEP         	: THDBEdit ;
    dbMntAgios		: THDBEdit ;
    dbMntTombe		: THDBEdit ;
    dbDevMontant     	: THDBEdit ;
    dbDevCtrValeur	: THDBEdit ;
    cbTauxPrecompte     : TDBCheckBox;
    cbValBO          	: TDBCheckBox ;
    cbTauxFixe       	: TDBCheckBox ;
    cbAgiosDeduit    	: TDBCheckBox ;
    cbAgiosPrecompte    : TDBCheckBox ;
    dbTauxResultat   	: THDBEdit ;
    dDateDepart      	: THDBEdit ;
    dDateFin         	: THDBEdit ;
    dDateOpe         	: THDBEdit ;
    cbBaseCalcul	: THDBValComboBox ;
    spNbJBanque		: THDBSpinEdit ;
    {$ENDIF}
    Num              	: string ;
    VBO              	: Boolean ;
    sens		: string ;
    AncienMontant	: Double ;
    ChargeOk            : Boolean;
    ValeurControl       : string;
    InChargement        : Boolean;

   //Evènements
    procedure CategorieOnChange      (Sender : TObject);
    procedure TransactionOnChange    (Sender : TObject);
    procedure dbDevMontantOnChange   (Sender : TObject);
    procedure dbDevCtrValeurOnChange (Sender : TObject);
    procedure dbMontantOnExit        (Sender : TObject);
    procedure dbMontantOnEnter       (Sender : TObject);
    procedure cbValBOOnClick         (Sender : TObject);
    procedure cbTauxFixeOnClick      (Sender : TObject);
    procedure dbTauxResultatOnExit   (Sender : TObject);
    procedure cbCompteOnExit   	     (Sender : TObject);
    procedure dDateDepartOnExit      (Sender : TObject);
    procedure dDateFinOnExit         (Sender : TObject);
    procedure dDateOpeOnExit         (Sender : TObject);
    procedure BCalcOnClick           (Sender : TObject);
    procedure BCalcTauxOnClick	     (Sender : TObject);
    procedure BCalcMntAgiosOnClick   (Sender : TObject);
    procedure spNbJBanqueOnChange    (Sender : TObject);
    procedure dbTransactionKeyPress  (Sender : TObject; var Key: Char);
    procedure OnApresShow            ;

    //gestion
    function  RempliCondition : Boolean ;
    procedure InitNumTransaction ;
    procedure ActiveLien(ok : boolean);
    procedure MAJDecimale;
  end ;

implementation

uses
   Commun, ParamSoc, Constantes;


{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_CourtsTermes(Dom, Fiche, Range, Lequel, Arguments : string ) ;
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments ); 
end ;

{---------------------------------------------------------------------------------------}
function TOM_CourtsTermes.RempliCondition : Boolean ;
{---------------------------------------------------------------------------------------}
Var
  FPl : TOB;
  Cpt : string;
Begin
  {En création sur le OnExit GetField renvoie vide...}
  Cpt := GetField('TCT_COMPTETR');
  {... On récupère alors la valeur directement dans la combo}
  if Cpt = '' then Cpt := GetControlText('TCT_COMPTETR');

  FPl := TOB.Create('CONDITIONFINPLAC', nil, -1);
  try
    Result := GetConditionFinPlac(FPl,GetField('TCT_TRANSAC'), '', Cpt, V_PGI.CodeSociete);
    if Result and
       (TrShowMessage(Ecran.Caption, 16, '', '') = mrYes) then begin
      with FPl do begin
        SetField('TCT_PRECOMPTE'    , GetString('TCF_AGIOSPRECOMPTE'));
        SetField('TCT_DEDUIT'       , GetString('TCF_AGIOSDEDUIT'));
        SetField('TCT_TAUXFIXE'     , GetString('TCF_TAUXFIXE'));
        SetField('TCT_BASECALCUL'   , GetInteger('TCF_BASECALCUL'));
        SetField('TCT_TAUXPRECOMPTE', GetString('TCF_TAUXPRECOMPTE'));
        SetField('TCT_VALTAUX'      , GetDouble('TCF_VALTAUX'));
        SetField('TCT_TAUX'         , GetString('TCF_TAUXVAR'));
        SetField('TCT_MAJORATION'   , GetDouble('TCF_VALMAJORATION'));
        SetField('TCT_MULTIPLE'     , GetDouble('TCF_VALMULTIPLE'));
        SetField('TCT_NBJOURBANQUE' , GetDouble('TCF_NBJOURBANQUE'));
      end;
    end;
  finally
    FreeAndNil(FPl);
  end;
  cbTauxFixeOnClick(cbTauxFixe);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CourtsTermes.ActiveLien( ok : boolean);
{---------------------------------------------------------------------------------------}
begin
   if ok then
   begin
      if dbCategorie     	<> nil then dbCategorie.OnChange    := CategorieOnChange ;
      if dbTransaction   	<> nil then dbTransaction.OnChange  := TransactionOnChange ;
      if dbTransaction   	<> nil then dbTransaction.OnKeyPress:= dbTransactionKeyPress;
      if dbMontant       	<> nil then dbMontant.OnExit        := dbMontantOnExit ;
      if dbMontant       	<> nil then dbMontant.OnEnter       := dbMontantOnEnter;
      if dbDevMontant    	<> nil then dbDevMontant.OnChange   := dbDevMontantOnChange ;
      if dbDevCtrValeur  	<> nil then dbDevCtrValeur.OnChange := dbDevCtrValeurOnChange ;
      if cbTauxFixe      	<> nil then cbTauxFixe.OnClick      := cbTauxFixeOnClick ;
      if cbValBO         	<> nil then cbValBO.OnClick         := cbValBOOnClick ;
      if dbTauxResultat  	<> nil then dbTauxResultat.OnExit   := dbTauxResultatOnExit ;
      if dDateFin        	<> nil then dDateFin.Onexit         := dDateFinOnExit ;
      if dDateDepart     	<> nil then dDateDepart.Onexit      := dDateDepartOnExit ;
      if dDateOpe        	<> nil then dDateOpe.Onexit         := dDateOpeOnExit ;
      if BCalc           	<> nil then BCalc.Onclick           := BCalcOnClick ;
      if BCalcTaux       	<> nil then BCalcTaux.Onclick       := BCalcTauxOnClick ;
      if BCalcMntAgios   	<> nil then BCalcMntAgios.Onclick   := BCalcMntAgiosOnClick ;
      if spNbJBanque	        <> nil then spNbJBanque.OnChange    := spNbJBanqueOnChange ;
      if cbCompte	 	<> nil then cbCompte.OnExit	    := cbCompteOnExit ;
   end
   else
   begin
      dbCategorie.OnChange    := Nil ;
      dbTransaction.OnChange  := Nil ;
      dbMontant.OnChange      := Nil ;
      dbDevMontant.OnChange   := Nil ;
      cbTauxFixe.OnClick      := Nil ;
      cbValBO.OnClick         := Nil ;
      dbTauxResultat.OnExit   := Nil ;
      dDateFin.Onexit         := Nil ;
      dDateDepart.Onexit      := Nil ;
      BCalc.Onclick           := Nil ;
      spNbJBanque.OnChange    := Nil ;
      cbContrePartie	      := Nil ;
      cbCompte	 	      := Nil ;
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 01/10/2001
Modifié le ... :   /  /
Description .. : Gestion de la catégorie de transaction
Mots clefs ... : TRANSACTION;CATEGORIE
*****************************************************************}
procedure TOM_CourtsTermes.CategorieOnChange(Sender : TObject) ;
begin
  if Length(dbCategorie.Text)>0 then
    dbTransaction.Plus:= 'TTR_CATTRANSAC="' + dbCategorie.Text + '"' ;
  if ds.state = dsInsert then
    InitNumTransaction ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Régis ROHAULT
Créé le ...... : 11/12/2001
Modifié le ... :   /  /
Description .. : Retourne la devise du compte sélectionné
Mots clefs ... : DEVISE;COMPTE
*****************************************************************}
procedure TOM_CourtsTermes.cbCompteOnExit(Sender : TObject) ;
var
  Cpt : string;
  Dev : string ;
begin
  {$IFDEF EAGLCLIENT}
  {Mis à False dans le OnLoadRecord}
  if not ChargeOk then begin
    ChargeOk := True;
    Exit;
  end;
  {$ENDIF}

  {En création sur le OnExit GetField renvoie vide...}
  Cpt := GetField('TCT_COMPTETR');
  {... On récupère alors la valeur directement dans la combo}
  if Cpt = '' then Cpt := GetControlText('TCT_COMPTETR');

  dev := RetDeviseCompte(Cpt) ;
  SetField('TCT_DEVMONTANT', dev);

  {Affichage des drapeaux et des libellés des devises}
//  dbDevMontantOnChange(dbDevMontant);
  
  {Récupération éventuelle de conditions paramétrées}
  RempliCondition;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 01/10/2001
Modifié le ... :   /  /
Description .. : Gestion du type de taux
Mots clefs ... : TAUX FIXE;TAUX VARIABLE
*****************************************************************}
PROCEDURE TOM_CourtsTermes.cbTauxFixeOnClick(Sender : TObject) ;
Var
  Bool : Boolean ;
BEGIN
  {On exécute ce code que si la fiche est chargée}
  if not Ecran.Active then Exit;

  Bool := cbTauxFixe.Checked ;
  if Bool then begin
    //si Taux Fixe
    SetField('TCT_TAUX','');
    SetField('TCT_MULTIPLE','100');
    SetField('TCT_MAJORATION','0');
    dbCodeTaux.Enabled := False ;
    dbMajoration.Enabled:=False;
    dbMultiple.Enabled:=False;
  end
  else begin
    // Si Taux Variable
    SetField('TCT_VALTAUX','1');
    dbcodetaux.Enabled   := True ;  //allume variable
    dbMajoration.Enabled := True;
    dbMultiple.Enabled   := True;
  end ;
  dbValTaux.Enabled := Bool;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 01/10/2001
Modifié le ... :   /  /
Description .. : gestion de la validation back office
Mots clefs ... : VALIDATION BACK OFFICE
*****************************************************************}
procedure TOM_CourtsTermes.cbValBOOnClick(Sender : TObject) ;
begin
  if cbValBO.Enabled then begin
    if cbValBO.Checked then begin
      setField('TCT_DATEVBO', Now);//DateTimeToStr(Now)
      Setfield('TCT_OPERATEURVBO',V_PGI.User);
    end
    else
    begin
      setField('TCT_DATEVBO', 2);// '01/01/1900'
      Setfield('TCT_OPERATEURVBO','');
    end ; //if
  end;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_CourtsTermes.spNbJBanqueOnChange(sender : TObject);
{---------------------------------------------------------------------------------------}
Var
  Nbj, NbjB : integer ;
Begin
  {On exécute ce code que si la fiche est chargée}
  if not Ecran.Active then Exit;

  Nbj := GetField('TCT_DATEFIN')-GetField('TCT_DATEDEPART');
  NbJB:= spNbJBanque.Value;
  SetField('TCT_NBJOUR', Nbj+NbJB);
End;

{---------------------------------------------------------------------------------------}
procedure TOM_CourtsTermes.dbTauxResultatOnExit(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
Begin
  if GetField('TCT_TAUXRESULTAT') = 0 then SetField('TCT_TAUXRESULTAT', 1);
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_CourtsTermes.dDateDepartOnExit(Sender : TObject);
{---------------------------------------------------------------------------------------}
{$IFDEF EAGLCLIENT}
var
  D, F : TDate;
  N    : Integer;
{$ENDIF}
begin
  {$IFDEF EAGLCLIENT}
  D := VarToDateTime(GetControlText('TCT_DATEDEPART'));
  F := VarToDateTime(GetControlText('TCT_DATEFIN'));
  N := StrToInt(VarToStr(GetControlText('TCT_NBJOURBANQUE')));

  if F < D then begin
    SetControlText('TCT_DATEFIN', DatetoStr(D));
    F := D;
  end;

  if GetControlText('TCT_DATEOPERATION') = DateToStr(2) then
    SetControlText('TCT_DATEOPERATION', DatetoStr(D));

  SetControlText('TCT_NBJOUR', FloatToStr(F - D + N)) ;
  {$ELSE}
  if GetField('TCT_DATEDEPART')>GetField('TCT_DATEFIN') then
    SetField('TCT_DATEFIN',GetField('TCT_DATEDEPART'));
  if VarToStr(GetField('TCT_DATEOPERATION')) = DateToStr(2) then
    SetField('TCT_DATEOPERATION',GetField('TCT_DATEDEPART'));
  SetField('TCT_NBJOUR',GetField('TCT_DATEFIN')-GetField('TCT_DATEDEPART')+ GetField('TCT_NBJOURBANQUE')) ;
  {$ENDIF}
  {On force un éventuel recalcul de la contre-valeur du montant}
  if GetControlText('TCT_DEVMONTANT') <> GetControlText('TCT_DEVCTRVALEUR') then begin
    ValeurControl := '0';
    dbMontantOnExit(dbMontant);
  end;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_CourtsTermes.dDateFinOnExit(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
{$IFDEF EAGLCLIENT}
var
  D, F : TDate;
  N    : Integer;
{$ENDIF}
Begin
  {$IFDEF EAGLCLIENT}
  D := VarToDateTime(GetControlText('TCT_DATEDEPART'));
  F := VarToDateTime(GetControlText('TCT_DATEFIN'));
  N := StrToInt(VarToStr(GetControlText('TCT_NBJOURBANQUE')));

  if F < D then begin
    SetControlText('TCT_DATEDEPART', DatetoStr(F));
    D := F;
  end;

  SetControlText('TCT_NBJOUR', FloatToStr(F - D + N));
  {$ELSE}
  if GetField('TCT_DATEDEPART')>GetField('TCT_DATEFIN') then
    SetField('TCT_DATEDEPART',GetField('TCT_DATEFIN'));
  SetField('TCT_NBJOUR',GetField('TCT_DATEFIN')-GetField('TCT_DATEDEPART')+ GetField('TCT_NBJOURBANQUE')) ;
  {$ENDIF}
End ;

{---------------------------------------------------------------------------------------}
procedure TOM_CourtsTermes.dDateOpeOnExit(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
Begin
  {$IFDEF EAGLCLIENT}
  if GetControlText('TCT_DATEDEPART') = DateToStr(2) then
    SetControlText('TCT_DATEDEPART', GetControlText('TCT_DATEOPERATION'));
  {$ELSE}
  if VarToStr(GetField('TCT_DATEDEPART')) = DateToStr(2) then
    SetField('TCT_DATEDEPART', GetField('TCT_DATEOPERATION'));
  {$ENDIF}
End ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 19/09/2001
Modifié le ... :   /  /
Description .. : Initilisation du numéro de dbTransaction
Mots clefs ... : NUMERO dbTransaction
*****************************************************************}
Procedure TOM_CourtsTermes.InitNumTransaction;
var
  str : string ;
begin
  if ((Length(dbCategorie.Text)>0) and (Length(dbTransaction.Text)>0) and (ds.State = dsInsert)) then begin
    if not InChargement then begin
       str := InitNumTransac(CodeModule, V_PGI.CodeSociete, dbTransaction.Text) ;
       Num := Copy(str,Length(str) - 6, 7) ;

      {JP 13/08/03 : Il ne faut surtout pas que le caractère '-' figure dans le numéro de
                     transaction :cf Commun.EcrireUneEcriture pour en comprendre la raison}
      while Pos('-', Str) > 0 do System.Delete(Str, Pos('-', Str), 1);

       SetField('TCT_NUMTRANSAC', str) ;
    end ;
  end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 01/10/2001
Modifié le ... :   /  /
Description .. : Gestion du la transaction
Mots clefs ... : TRANSACTION
*****************************************************************}
PROCEDURE TOM_CourtsTermes.TransactionOnChange(Sender : TObject) ;
BEGIN
  inherited;
  InitNumTransaction ;
END;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 01/10/2001
Modifié le ... : 01/10/2001
Description .. : Gestion de la devise de transaction
Mots clefs ... : DEVISE TRANSACTION
*****************************************************************}
procedure TOM_CourtsTermes.dbDevMontantOnChange(Sender : TObject);
begin
  inherited ;
  MAJDecimale;
  AssignDrapeau(Timage(GetControl('IDEV1')), GetField('TCT_DEVMONTANT'));
  AssignDrapeau(Timage(GetControl('IDEV3')), GetField('TCT_DEVMONTANT'));
  lDev.Caption:=GetField('TCT_DEVMONTANT');
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_CourtsTermes.dbDevCtrValeurOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
Begin
  inherited;
  AssignDrapeau(Timage(GetControl('IDEV2')), GetField('TCT_DEVCTRVALEUR'));
  MAJDecimale;
End ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 01/10/2001
Modifié le ... :   /  /
Description .. : Gestion du montant de la transaction
Mots clefs ... : MONTANT INITIAL;MONTANT TRANSACTION
*****************************************************************}
procedure TOM_CourtsTermes.dbMontantOnExit(Sender : TObject);
var
  aDte : TDateTime ;
  Mnt  : Double;
  Dev1 : string;
  Dev2 : string;
  Val  : Double ;
begin
  inherited ;
  if ValeurControl = FloatToStr(Valeur(GetControlText('TCT_MONTANT'))) then Exit;

  aDte := StrToDate(GetControlText('TCT_DATEDEPART'));
  {Si la date n'a pas été saisie}
  if aDte < 3 then aDte := V_PGI.DateEntree;
  mnt  := Valeur(GetControlText('TCT_MONTANT'));
  Dev1 := GetControlText('TCT_DEVMONTANT');
  Dev2 := GetControlText('TCT_DEVCTRVALEUR');
  Val := RetContreValeur(aDte, Mnt, Dev1, Dev2) ;
  SetField('TCT_MONTANTCTRVAL', FloatToStr(Val));
//  dbMontantCtrVal.Text  := FloatToStr(Val)
end ;

{JP 29/10/03 : Pour éviter que le traitement du OnExit soit systématiquement éxécuté
{---------------------------------------------------------------------------------------}
procedure TOM_CourtsTermes.dbMontantOnEnter(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  ValeurControl := FloatToStr(Valeur(GetControlText('TCT_MONTANT')));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Régis ROHAULT
Créé le ...... : 11/12/2001
Modifié le ... : 11/12/2001
Description .. : Mise à jour des décimales des montants de la transaction
Suite ........ : par rapport à la devise du compte
Mots clefs ... : DEVISE;DECIMALES
*****************************************************************}
procedure TOM_CourtsTermes.MAJDecimale;
var
  Mask : string;
begin
  dbMontantCtrVal.DisplayFormat := strFMask(CalcDecimaleDevise(GetField('TCT_DEVCTRVALEUR')),'',true);
  Mask := strFMask(CalcDecimaleDevise(GetField('TCT_DEVMONTANT')),'',true);
  dbMontant.DisplayFormat:= Mask;
  dbMntMEP.DisplayFormat:=  Mask;
  dbMntAgios.DisplayFormat:=Mask;
  dbMntTombe.DisplayFormat:=Mask;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Régis ROHAULT
Créé le ...... : 11/12/2001
Modifié le ... :   /  /
Description .. : Calcul automatique de l'ensemble du résultat
Mots clefs ... : CALCUL;RESULTAT
*****************************************************************}
PROCEDURE TOM_CourtsTermes.BCalcOnClick(Sender : TObject);
Var
  mnt,val,d : double;
  NbjBase : integer ;
  Base 	: integer ;
  dd 	: TDateTime ;
  Taux : double;
  Mul : double ;
  Maj : double ;
begin
  if cbBaseCalcul.ItemIndex < 0 then
    cbBaseCalcul.ItemIndex := 0;
  Base :=cbBaseCalcul.ItemIndex;
  dd := GetField('TCT_DATEDEPART') ;
  if dd < 0 then dd := EncodeDate(1900,01,01);
  NbJBase := CalcNbJourParAnBase(dd, Base) ;

  //MAJ des Montant
  Setfield('TCT_MNTMEP', dbMontant.Text);
  Setfield('TCT_MNTTOMBE', dbMontant.Text);

    //Calcul du Taux
  if cbTauxFixe.checked then begin
    val := GetField('TCT_VALTAUX');
    //Par défaut c'est PostCompté
    d := val ;
    if cbTauxPrecompte.checked then d := CalcTauxPrePost(false , val , GetField('TCT_NBJOUR'), NbJBase) ;
    SetField('TCT_TAUXRESULTAT', d);  // on affiche toujours le taux postcompté
  end
  else begin
    GetValTauxRef(GetControlText('TCT_TAUX'), GetField('TCT_DATECREATION'), Taux );
    Maj := getField('TCT_MAJORATION') ;
    Mul := getField('TCT_MULTIPLE') ;
    taux := (Taux+Maj)* (Mul/100) ;
    SetField('TCT_TAUXRESULTAT', FloatTosTr(Taux));
  end ;

  //Calcul les agios/intérets par rapport à un taux résultat postcompté !!!
  Mnt :=CalcMontantAgios(GetField('TCT_MONTANT'), GetField('TCT_TAUXRESULTAT'),
  NbjBase, GetField('TCT_NBJOUR'), cbTauxPrecompte.Checked , cbAgiosPrecompte.checked) ;
  SetField('TCT_MNTAGIOS', Mnt);

  //Calcul par rapport au check AgiosDeduit
  if cbAgiosPrecompte.checked and cbAgiosDeduit.Checked then
    SetField('TCT_MNTMEP', GetField('TCT_MONTANT')-GetField('TCT_MNTAGIOS'))
  else
    SetField('TCT_MNTMEP', GetField('TCT_MONTANT'));
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Régis ROHAULT
Créé le ...... : 11/12/2001
Modifié le ... :   /  /
Description .. : calcul du taux lors de la modification des Agios
Mots clefs ... : AGIOS;TAUX
*****************************************************************}
PROCEDURE TOM_CourtsTermes.BCalcTauxOnClick(Sender : TObject);
var base : integer ;
	Mnt : Double ;
Begin
  base := CalcNbJourParAnBase(GetField('TCT_DATEDEPART'),GetField('TCT_BASECALCUL')) ;
  Mnt := CalcMontantAgios( GetField('TCT_MONTANT'),GetField('TCT_TAUXRESULTAT'),base,GetField('TCT_NBJOUR'), cbTauxPreCompte.checked, cbAgiosPrecompte.Checked) ;
  SetField( 'TCT_MNTAGIOS',Mnt);
End ;

{***********A.G.L.***********************************************
Auteur  ...... : Régis ROHAULT
Créé le ...... : 11/12/2001
Modifié le ... :   /  /
Description .. : Calcul du montant des Agios lors de la modification du Taux
Mots clefs ... : TAUX;AGIOS
*****************************************************************}
procedure TOM_CourtsTermes.BCalcMntAgiosOnClick(Sender : TObject);
var
  base : integer ;
  Taux : Double ;
Begin
  base := CalcNbJourParAnBase(GetField('TCT_DATEDEPART'),GetField('TCT_BASECALCUL')) ;
  Taux := CalcTauxAgios( GetField('TCT_MNTAGIOS'),base,GetField('TCT_MONTANT'),GetField('TCT_NBJOUR'),True ,not cbAgiosPrecompte.Checked) ;
  SetField( 'TCT_TAUXRESULTAT',Taux);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CourtsTermes.dbTransactionKeyPress(Sender : TObject; var Key: Char);
{---------------------------------------------------------------------------------------}
begin
  ValidCodeOnKeyPress(Key);
  {Il ne faut surtout pas que le caractère '-' figure dans le numéro de transaction :
   cf Commun.EcrireUneEcriture pour en comprendre la raison}
   if Key = '-' then key := #0;
end;

{*********************************************************************
 **                             TOM                                 **
 *********************************************************************}

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 17/09/2001
Modifié le ... :   /  /
Description .. : Initialisation du numéro de dbTransaction
Mots clefs ... : dbTransaction, NUMÉRO
*****************************************************************}
procedure TOM_CourtsTermes.OnNewRecord ;
BEGIN
  if ds.state = dsinsert then begin
    {La contre-valeur est dans la devise pivot de la compta}
    SetField('TCT_DEVCTRVALEUR', V_PGI.DevisePivot);
    {Chargement du drapeau}
    AssignDrapeau(Timage(GetControl('IDEV2')), GetField('TCT_DEVCTRVALEUR'));
    {Initialisation avce le paramsoc}
    SetField('TCT_SOCIETE', V_PGI.CodeSociete);
    SetField('TCT_NODOSSIER', V_PGI.NoDossier); {29/05/06 : FQ 10360}

    SetField('TCT_TAUXFIXE','X') ;
    SetField('TCT_TAUXFIXE','X') ;
    SetField('TCT_TAUXPRECOMPTE','-') ;
    SetField('TCT_VALBO','-');
    SetField('TCT_BASECALCUL', '0');
    setField('TCT_DATECREATION', V_PGI.DateEntree);
    setField('TCT_OPERATEURCRE',V_PGI.User) ;
    setField('TCT_DENOUEUR', '') ;
    setField('TCT_DATEDENOUAGE', 2) ;
  end ;
  inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_CourtsTermes.OnAfterUpdateRecord ;
{---------------------------------------------------------------------------------------}
var
  str : string ;
begin
  inherited;
  {Mise à jour du numéro de transaction dans la table COMPTEURTRESO}
  if TFFiche(Ecran).TypeAction = tacreat then begin
    str := GetNum(CodeModule, V_PGI.CodeSociete, dbTransaction.Text);
    if str = Num then  // si c'est le même pas de pb
      Setnum(CodeModule, V_PGI.CodeSociete, dbTransaction.text,num)
    else begin
      InitNumTransaction; // sinon on recréer un num de dbTransaction
      SetNum(CodeModule, V_PGI.CodeSociete, dbTransaction.text,str);
    end ;
  end;
  Ecran.Close;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CourtsTermes.OnUpdateRecord ;
{---------------------------------------------------------------------------------------}
begin
  inherited ;
  if TFFiche(Ecran).TypeAction = taModif then begin
    setField('TCT_DATEMODIFI', V_PGI.DateEntree);
    setField('TCT_OPERATEURMOD', V_PGI.User);
  end ;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CourtsTermes.OnLoadRecord ;
{---------------------------------------------------------------------------------------}
begin
  inherited ;
  {Chargement du drapeau}
  AssignDrapeau(Timage(GetControl('IDEV2')), GetField('TCT_DEVCTRVALEUR'));
  {21/01/05 : On garde le numéro Calculé}
  SetControlEnabled('TCT_NUMTRANSAC', False);

  VBO := GetField('TCT_VALBO') = 'X';
  ancienMontant := GetField('TCT_MNTMEP');
  cbTauxFixeOnClick(Nil);
  {Pour ne pas exécuter le cbCompteOnExit lors du chargement d'un enregistrement}
  ChargeOk := False;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CourtsTermes.OnArgument(Arguments : string) ;
{---------------------------------------------------------------------------------------}
Var
  Action     : string;
  CatTransac : string ;
begin
  InChargement := True;
  inherited;
  Ecran.HelpContext := 150;
  Action       := ReadTokenSt(Arguments);
  Sens         := ReadTokenSt(Arguments);
  CatTransac   := ReadTokenSt(Arguments);

  SetControlCaption('LSOC', V_PGI.NomSociete);
  {$IFDEF EAGLCLIENT}
  dbCategorie      := THEdit(GetControl('TCT_CATTRANSAC')) ;
  dbTransaction    := THEdit(GetControl('TCT_TRANSAC')) ;
  cbValBO          := TCheckBox(GetControl('TCT_VALBO'));
  cbTauxFixe       := TCheckBox(GetControl('TCT_TAUXFIXE')) ;
  dbValTaux        := THEdit(GetControl('TCT_VALTAUX')) ;
  dbCodeTaux       := THEdit(GetControl('TCT_TAUX')) ;
  dbMajoration     := THEdit(GetControl('TCT_MAJORATION')) ; ;
  dbMultiple       := THEdit(GetControl('TCT_MULTIPLE')) ;
  dbMontant        := THEdit(GetControl('TCT_MONTANT')) ;
  dbMontantCtrVal  := THEdit(GetControl('TCT_MONTANTCTRVAL')) ;
  dbMntMEP         := THEdit(GetControl('TCT_MNTMEP')) ;
  dbMntAgios	   := THEdit(GetControl('TCT_MNTAGIOS')) ;
  dbMntTombe	   := THEdit(GetControl('TCT_MNTTOMBE')) ;
  dbDevMontant     := THEdit(GetControl('TCT_DEVMONTANT')) ;
  dbDevCtrValeur   := THEdit(GetControl('TCT_DEVCTRVALEUR')) ;
  cbTauxPrecompte  := TCheckBox(GetControl('TCT_TAUXPRECOMPTE'));
  cbAgiosDeduit    := TCheckBox(GetControl('TCT_DEDUIT'));
  cbAgiosPrecompte := TCheckBox(GetControl('TCT_PRECOMPTE'));
  dbTauxResultat   := THEdit(GetControl('TCT_TAUXRESULTAT')) ;
  dDateDepart      := THEdit(GetControl('TCT_DATEDEPART')) ;
  dDateFin         := THEdit(GetControl('TCT_DATEFIN')) ;
  dDateOpe         := THEdit(GetControl('TCT_DATEOPERATION')) ;
  spNbJBanque	   := THSpinEdit(GetControl('TCT_NBJOURBANQUE'));
  cbBaseCalcul	   := THValComboBox(GetControl('TCT_BASECALCUL')) ;
  cbCompte	   := THEdit(GetControl('TCT_COMPTETR')) ;
  cbContrePartie   := THEdit(GetControl('TCT_CONTREPARTIETR')) ;
  {$ELSE}
  dbCategorie      := THDBEdit(GetControl('TCT_CATTRANSAC')) ;
  dbTransaction    := THDBEdit(GetControl('TCT_TRANSAC')) ;
  cbValBO          := TDBCheckBox(GetControl('TCT_VALBO'));
  cbTauxFixe       := TDBCheckBox(GetControl('TCT_TAUXFIXE')) ;
  dbValTaux        := THDBEdit(GetControl('TCT_VALTAUX')) ;
  dbCodeTaux       := THDBEdit(GetControl('TCT_TAUX')) ;
  dbMajoration     := THDBEdit(GetControl('TCT_MAJORATION')) ; ;
  dbMultiple       := THDBEdit(GetControl('TCT_MULTIPLE')) ;

  dbMontant        := THDBEdit(GetControl('TCT_MONTANT')) ;
  dbMontantCtrVal  := THDBEdit(GetControl('TCT_MONTANTCTRVAL')) ;
  dbMntMEP         := THDBEdit(GetControl('TCT_MNTMEP')) ;
  dbMntAgios	   := THDBEdit(GetControl('TCT_MNTAGIOS')) ;
  dbMntTombe	   := THDBEdit(GetControl('TCT_MNTTOMBE')) ;

  dbDevMontant     := THDBEdit(GetControl('TCT_DEVMONTANT')) ;
  dbDevCtrValeur   := THDBEdit(GetControl('TCT_DEVCTRVALEUR')) ;
  cbTauxPrecompte  := TDBCheckBox(GetControl('TCT_TAUXPRECOMPTE'));
  cbAgiosDeduit    := TDBCheckBox(GetControl('TCT_DEDUIT'));
  cbAgiosPrecompte := TDBCheckBox(GetControl('TCT_PRECOMPTE'));
  dbTauxResultat   := THDBEdit(GetControl('TCT_TAUXRESULTAT')) ;
  dDateDepart      := THDBEdit(GetControl('TCT_DATEDEPART')) ;
  dDateFin         := THDBEdit(GetControl('TCT_DATEFIN')) ;
  dDateOpe         := THDBEdit(GetControl('TCT_DATEOPERATION')) ;
  spNbJBanque	   := THDBSpinEdit(GetControl('TCT_NBJOURBANQUE'));
  cbBaseCalcul	   := THDBValComboBox(GetControl('TCT_BASECALCUL')) ;
  cbCompte	   := THDBEdit(GetControl('TCT_COMPTETR')) ;
  cbContrePartie   := THDBEdit(GetControl('TCT_CONTREPARTIETR')) ;
  {$ENDIF}

  lCodeTaux        := THLabel(GetControl('TTCT_TAUX')) ; ;
  lPlus            := THLabel(GetControl('PLUS')) ;
  lMultiple        := THLabel(GetControl('MULTIPLE')) ;
  lDev             := THLabel(GetControl('DEV')) ;
  BCalc            := TBitBtn(GetControl('BCALC')) ;
  BCalcTaux	   := TBitBtn(GetControl('BCALCTAUX')) ;
  BCalcMntAgios	   := TBitBtn(GetControl('BCALCMNTAGIOS')) ;

  {JP 09/03/04 : Nouveau champ dans la version 4 de la table CATTRANSAC : TCA_TYPETRANSAC}
  dbCategorie.Plus := 'TCA_SENS="' + Sens + '" AND TCA_TYPETRANSAC = "' + CatTransac + '"';

       if action='ACTION=CREATION'     then TFFiche(Ecran).TypeAction := tacreat
  else if action='ACTION=MODIFICATION' then TFFiche(Ecran).TypeAction := taModif
  else if action='ACTION=CONSULTATION' then TFFiche(Ecran).TypeAction := taConsult;

  {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
  cbCompte.Plus := FiltreBanqueCp(cbCompte.DataType, '', '');
  
  ActiveLien(True);
  TFFiche(Ecran).OnAfterFormShow := OnApresShow;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_CourtsTermes.OnApresShow;
{---------------------------------------------------------------------------------------}
begin
  InChargement := False;
end;

{18/09/06 : Gestion du multi-sociétés
{---------------------------------------------------------------------------------------}
procedure TOM_CourtsTermes.OnChangeField(F : TField);
{---------------------------------------------------------------------------------------}
var
  R : TDblResult;
begin
  inherited;
  if IsTresoMultiSoc and (F.FieldName = 'TCT_COMPTETR') then begin
    R := GetInfosSocFromBQ(F.AsString, True);
    SetField('TCT_NODOSSIER', R.RE);
    SetField('TCT_SOCIETE', R.RC);
    SetControlCaption('LSOC', R.RT);
  end;
end;

Initialization
  RegisterClasses([TOM_CourtsTermes]) ;
END .

