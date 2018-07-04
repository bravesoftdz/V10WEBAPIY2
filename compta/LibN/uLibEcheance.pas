unit uLibEcheance;


interface

uses
  SysUtils,
  HEnt1,
  Ent1,
  SaisUtil,        // RDevise
  uLibEcriture,    // Pour TInfoErriture
  ParamSoc,        // GetParamSocSecur
  Echeance,        // Pout EcheArrondi
  uTob ;


Type

TMultiEche = Class(TOB)
  //////////////////////////////////////////////////////////////////////////////////////////
  private
    FStModeInit     : String ;
    FStModeFinal    : String ;
    FBoModifTVA     : Boolean ;
    FTobMR          : TOB ;
    FTobEcr         : TOB ;
    FInfoEcr        : TInfoEcriture ;
    FBoModif        : Boolean ;
    FAction         : TActionFiche ;

    function  GetNumLigne    : Integer ;
    function  GetNbEche      : Integer ;
    function  GetDevise      : RDevise ;

  //////////////////////////////////////////////////////////////////////////////////////////
  Public

    // ======================
    // ===== PROPRIETES =====
    // ======================
    property Devise       : RDevise       read  GetDevise ;
    property ModeInit     : String        read  FStModeInit   write  FStModeInit ;
    property ModeFinal    : String        read  FStModeFinal  write  FStModeFinal ;
    property ModifTVA     : Boolean       read  FBoModifTVA   write  FBoModifTVA;
    property ModifEnCours : Boolean       read  FBoModif      write  FBoModif ;
    property TobMR        : TOB           read  FTobMR ;
    property TobEcr       : TOB           read  FTobEcr ;
    property InfoEcr      : TInfoEcriture read  FInfoEcr ;
    property Action       : TActionFiche  read  FAction       write FAction ;
    property NumLigne     : Integer       read  GetNumLigne ;
    property NbEche       : Integer       read  GetNbEche ;

    // ====================
    // ===== METHODES =====
    // ====================

    // Constructeur / Destructeur
    constructor Create( vTobParent : TOB ; vTobEcr : TOB ; vInfoEcr : TInfoEcriture ; vAction : TActionFiche ) ;
    destructor  Destroy ; override ;


    // Indicateurs
    function  EstDebit       : Boolean ;
    function  EstTvaLoc      : Boolean ;
    function  GetMontantEcr  : Double ;
    function  GetDateFact    : TDateTime ;
    function  GetDateBL      : TDateTime ;
    function  GetDateFactExt : TDateTime ;
    function  EstLettre      ( vNumEche : Integer ) : Boolean ;
    function  EstModifiable  ( vNumEche : Integer = -1 ) : Boolean ;
    function  GetMontantEche ( vNumEche : Integer ) : Double ;
    function  GetTotalEche   : Double ;
    function  EstOkMultiEche : Boolean ;
    procedure ChargeMR ;
    procedure ChargeInfoEcr ;

    // Gestion des lignes
    procedure Trier ;
    function  IsOut          ( vNumEche : Integer ) : Boolean ;
    procedure AddEche        ( vTobEcr : TOB ) ;
    function  NewEche        ( vNumEche : Integer = 0 ) : TOB ;
    function  GetEche        ( vNumEche : Integer ) : TOB ;

    // Calculs de l'échéance
    Function  GetModeInit : String ;
    Function  CalculModeFinal  ( vStModeForce : String = '' ) : TOB ;
    Function  GetDateDepart : TDateTime ;
    procedure CalculEche( vStModeForce : String = '' ) ;
    procedure ProratiseEche ( vMontant : Double ) ;
    procedure UpdateEche     ( vNumEche : Integer ; vModePaie : String ; vMontant : Double ; vDateEche : TDateTime ) ;
    function  GereArrondi : Boolean;
    procedure GerePoucentage ;
    procedure CopieEcheances     ( vSource : TMultiEche ) ;

    // Gestions des cumuls / synchronisation avec données de l'écriture
    procedure CumuleTobEcr ;
    procedure SynchroniseEcheances ;
    procedure SynchroniseEcr ;
    procedure Renumerote ;
    Procedure PutValueAll( vChamps : String ; vValeur : Variant ) ;
    Procedure PutTobEche ( vNumEche : Integer ; vTobSource : TOB ) ;

    // Calcul des champs TVa
    Procedure SetTvaEnc( vTvaEnc1, vTvaEnc2, vTvaEnc3, vTvaEnc4, vTvaDebit : Double ) ;
    Procedure InitTvaEnc ;

    // Compatibilité T_MODEREGL <==> TMULTIECHE
    Function  GetModeRegl  : T_MODEREGL ;
    procedure SetModeRegl  ( vModeRegl : T_MODEREGL ) ;

  end;

// Fonctions de création des échéances
Function  CCreerMultiEche ( vParent : TOB ; vTobEcr : TOB ; vInfoEcr : TInfoEcriture ; vAction : TActionFiche ; vBoAvecInit : Boolean = False ) : TMultiEche ;


implementation


Function  CCreerMultiEche ( vParent : TOB ; vTobEcr : TOB ; vInfoEcr : TInfoEcriture ; vAction : TActionFiche ; vBoAvecInit : Boolean = False ) : TMultiEche ;
begin
  result := TMultiEche.Create( vParent, vTobEcr, vInfoEcr, vAction ) ;
  if vBoAvecInit then
    begin
    // Initialisation pour les compte divers lettrables spécifiques FQ 16136 SBO 23/08/2005
    if (vTobEcr.GetString('E_GENERAL') <> '') and vInfoEcr.LoadCompte( vTobEcr.GetString('E_GENERAL') ) and
       ( vInfoEcr.Compte.GetValue('G_NATUREGENE') = 'DIV') and
       vInfoEcr.Compte.IsLettrable
       then result.CalculEche( GetParamSocSecur('SO_GCMODEREGLEDEFAUT','') )
       else result.CalculEche ;
    end ;

  if result.NbEche = 0 then
    result.NewEche ;

end ;



{ TMultiEche }


constructor TMultiEche.Create( vTobParent : TOB ; vTobEcr : TOB ; vInfoEcr : TInfoEcriture ; vAction : TActionFiche ) ;
begin
  inherited Create('V_ECHEANCES', vTobParent, -1) ;

  FTobEcr       := vTobEcr ;
  FInfoEcr      := vInfoEcr ;

  FStModeInit   := GetModeInit ;
  FStModeFinal  := '' ;

  FBoModif      := False ;
  FBoModifTVA   := False ;

  FAction       := vAction ;

end;

destructor TMultiEche.Destroy;
begin
  if Assigned( FTobMR ) then
    FreeAndNil( FTobMR ) ;
  inherited;
end;

function TMultiEche.EstDebit: Boolean;
begin
  result := TobEcr.GetDouble('E_DEBITDEV')<>0 ;
end;

function TMultiEche.GetDateFact: TDateTime;
begin
  result := TobEcr.GetDateTime('E_DATECOMPTABLE') ;
end;

function TMultiEche.GetNumLigne: Integer;
begin
  result := TobEcr.GetInteger('E_NUMLIGNE') ;
end;

function TMultiEche.GetMontantEcr: Double;
begin
  result := TobEcr.GetDouble('E_CREDITDEV') + TobEcr.GetDouble('E_DEBITDEV') ;
end;

function TMultiEche.GetNbEche: Integer;
begin
  result := Detail.Count ;
end;

procedure TMultiEche.CalculEche( vStModeForce : String = '' ) ;
var lDtDepart      : TDateTime ;
    lInEche        : Integer ;
    lInNbEche      : Integer ;
    lTaux          : Double ;
    lStModePaie    : String ;
    lTobEche       : Tob ;
    lMontant       : Double ;
    lTotal         : Double ;
    lMontantEcr    : Double ;
begin

  // ==========================================================
  // ==== Détermination du mode de règlement à appliquer : ====
  // ==========================================================
  if not Assigned( FTobMR ) or
     ( FTobMR.GetString('MR_MODEREGLE') <> vStModeForce ) then
    begin
    if Assigned( FTobMR ) then FreeAndNil( FTobMR ) ;
    if vStModeForce <> '' then
      begin
      FStModeFinal := vStModeForce ;
      ChargeMR ;
      end
    else FTobMR := CalculModeFinal ;
    end ;
  if FTobMR = nil then Exit ;

  // ============================================
  // ==== Détermination de la date de départ ====
  // ============================================
  lDtDepart := GetDateDepart ;

  // ====================================
  // ==== Construction des échéances ====
  // ====================================
  ClearDetail ;
  lMontant    := 0 ;
  lTotal      := 0 ;
  if EstOkMultiEche
    then lInNbEche := TobMR.GetInteger('MR_NOMBREECHEANCE')
    else lInNbEche := 1 ;

  lMontantEcr := GetMontantEcr ;

  For lInEche := 1 to lInNbEche do
    begin
    // Création de l'objet
    lTobEche := NewEche( lInEche ) ;

    // Détermination Mode et Taux
    lStModePaie := FTobMR.GetString('MR_MP' + IntToStr(lInEche)) ;
    lTaux       := FTobMR.GetDouble('MR_TAUX' + IntToStr(lInEche)) ;

    // Traitement ligne / dernière ligne
    if lInEche < lInNbEche then
      begin
      if lTaux <> 0 then
        begin
        lMontant := Arrondi(lMontantEcr * lTaux / 100.0, Devise.Decimale );
        lTotal   := Arrondi(lTotal + lMontant, Devise.Decimale );
        end ;
      end
    else
      begin
      lMontant := Arrondi( lMontantEcr - lTotal, Devise.Decimale );
      end;

    // affectation des montants
    if EstDebit
      then CSetMontants( lTobEche,   lMontant,    0,           Devise,   True )
      else CSetMontants( lTobEche,   0,           lMontant,    Devise,   True ) ;

    // Affectation de la date d'échéance
    lDtDepart := EcheArrondie( lDtDepart, FTobMR.GetString('MR_ARRONDIJOUR'), 0, 0);
    lTobEche.PutValue( 'E_DATEECHEANCE',    lDtDepart  ) ;
    lTobEche.PutValue( 'E_ORIGINEPAIEMENT', lDtDepart  ) ;

    // Affectation du mode de paiement
    lTobEche.PutValue( 'E_MODEPAIE',      lStModePaie  ) ;
    lTobEche.PutValue( 'E_CODEACCEPT',    MPTOACC ( lStModePaie ) ) ;

    // Calcul de la prochaine date
    NextEche( lDtDepart, FTobMR.GetString('MR_SEPAREPAR') );

    end ;

  FBoModif := True ;

end;

Function TMultiEche.CalculModeFinal ( vStModeForce : String ) : TOB ;

  function _DetermineMR( vStModeEnCours : string ; vMontantEcr : Double ) : TOB ;
  var lTobMR : TOB ;
  begin
    lTobMR := Tob.Create('MODEREGL', nil, -1) ;
    lTobMR.PutValue('MR_MODEREGLE', vStModeEnCours ) ;
    if lTobMR.LoadDB then
      begin
      if ( vMontantEcr >= lTobMR.GetDouble('MR_MONTANTMIN') ) 
         or ( lTobMR.GetString('MR_REMPLACEMIN') = '' )
         or ( lTobMR.GetString('MR_REMPLACEMIN') = vStModeEnCours )
      then result := lTobMR
      else
        begin
        result := _DetermineMR( lTobMR.GetString('MR_REMPLACEMIN'), vMontantEcr ) ;
        if result = nil
          then result := lTobMR
          else FreeAndNil( lTobMR ) ;
        end ;
      end
    else
      begin
      FreeAndNil( lTobMR ) ;
      result := nil ;
      end ;
  end ;

var lMontantEcr : Double ;
begin

  result       := nil ;

  // Montant à répartir
  lMontantEcr := GetMontantEcr ;
  if lMontantEcr = 0 then Exit ;

  // Mode de règlement paramètré dans les fiches auxi / généraux OU Forcé
  if vStModeForce <> ''
    then FStModeInit := vStModeForce
    else FStModeInit := GetModeInit ;
  if FStModeInit = '' then Exit ;

  result := _DetermineMR( FStModeInit , lMontantEcr ) ;
  FStModeFinal := result.GetValue('MR_MODEREGLE') ;

end;

Function TMultiEche.GetModeInit : String ;
begin
  result := '' ;
  if InfoEcr.LoadCompte( TobEcr.GetValue('E_GENERAL') ) then
    if InfoEcr.Compte.IsCollectif then
      begin
      if InfoEcr.LoadAux( FTobEcr.GetValue('E_AUXILIAIRE') ) then
        result  := InfoEcr.Aux_GetValue('T_MODEREGLE') ;
      end
    else
      result    := InfoEcr.Compte_GetValue('G_MODEREGLE') ;
end;

function TMultiEche.GetDateDepart: TDateTime;
var lStAPartirDe : String ;
begin

  result := FTobEcr.GetDateTime('E_DATECOMPTABLE') ;

  lStAPartirDe := TobMR.GetValue('MR_APARTIRDE') ;

  // Choix du départ
  if lStAPartirDe = '' then Exit ;
  if lStAPartirDe = 'FIN' then
    result := FinDeMois( result )
  else if lStAPartirDe = 'DEB' then
    result := DebutDeMois( result )
  else if lStAPartirDe = 'FAC' then
    begin
    if TobEcr.GetDateTime('E_DATEREFEXTERNE') > IDate1900 then
      result := TobEcr.GetDateTime('E_DATEREFEXTERNE') ;
    end
  else if lStAPartirDe = 'FAF' then
    begin
    if TobEcr.GetDateTime('E_DATEREFEXTERNE') > IDate1900 then
      result := TobEcr.GetDateTime('E_DATEREFEXTERNE') ;
    result := FinDeMois(result);
    end;

  result := result + TobMR.GetInteger('MR_PLUSJOUR') ;

end;

function TMultiEche.NewEche ( vNumEche : Integer ) : TOB ;
var lIndex : Integer ;
begin
  if (vNumEche < 1) or (vNumEche > NbEche )
    then lIndex := -1
    else lIndex := vNumEche - 1 ;

  result := Tob.Create('ECRITURE', self, lIndex ) ;
  result.Dupliquer( TobEcr, False, True, True ) ;

  if lIndex < 0
    then result.PutValue('E_NUMECHE', NbEche )
    else result.PutValue('E_NUMECHE', vNumEche ) ;

end;

procedure TMultiEche.ProratiseEche(vMontant: Double);
Var lTotal       : Double ;
    lTaux        : Double ;
    lInEche      : Integer ;
    lMontantEche : Double ;
    lTobEche     : TOB ;
begin

  lTotal := GetTotalEche ;
  If lTotal = 0 then Exit ;

// Lek 11/05/06  lTaux := lTotal / vMontant ;
  lTaux := vMontant / lTotal ;
  if lTaux = 1 then Exit ;

  lTotal    := 0 ;

  // Recalcul des montants des échéances
  for lInEche := 1 to NbEche do
    begin
    lTobEche := Detail[ lInEche - 1 ] ;

    lMontantEche := lTobEche.GetDouble('E_DEBITDEV') + lTobEche.GetDouble('E_CREDITDEV') ;
    lMontantEche := Arrondi( lTaux * lMontantEche , Devise.Decimale ) ;

    if EstDebit
      then CSetMontants( lTobEche, lMontantEche, 0, Devise, True )
      else CSetMontants( lTobEche, 0, lMontantEche, Devise, True ) ;

    // #TVAENC
    if ((VH^.OuiTvaEnc) and (ModifTVA) and (lTaux<>1)) then
      begin
      lTobEche.PutValue('E_ECHEENC1',  Arrondi( lTaux * lTobEche.GetDouble('E_ECHEENC1') ,  V_PGI.OkDecV ) )  ;
      lTobEche.PutValue('E_ECHEENC2',  Arrondi( lTaux * lTobEche.GetDouble('E_ECHEENC2') ,  V_PGI.OkDecV ) )  ;
      lTobEche.PutValue('E_ECHEENC3',  Arrondi( lTaux * lTobEche.GetDouble('E_ECHEENC3') ,  V_PGI.OkDecV ) )  ;
      lTobEche.PutValue('E_ECHEENC4',  Arrondi( lTaux * lTobEche.GetDouble('E_ECHEENC4') ,  V_PGI.OkDecV ) )  ;
      lTobEche.PutValue('E_ECHEDEBIT', Arrondi( lTaux * lTobEche.GetDouble('E_ECHEDEBIT') , V_PGI.OkDecV ) )  ;
      end ;

    lTotal := lTotal + lMontantEche ;

    end ;

  // Gestion des différences d'arrondi
  if lTotal <> vMontant then
    begin
    lTobEche := Detail[ NbEche - 1 ] ;
    lMontantEche := lTobEche.GetDouble('E_DEBITDEV') + lTobEche.GetDouble('E_CREDITDEV') ;
    lMontantEche := lMontantEche + vMontant - lTotal ;
    if EstDebit
      then CSetMontants( lTobEche, lMontantEche, 0, Devise, True )
      else CSetMontants( lTobEche, 0, lMontantEche, Devise, True ) ;
    end ;

end;

procedure TMultiEche.ChargeMR;
begin
  FTobMR := Tob.Create('MODEREGL', nil, -1) ;
  FTobMR.PutValue('MR_MODEREGLE', ModeFinal ) ;
  FTobMR.LoadDB ;
end;

procedure TMultiEche.AddEche(vTobEcr: TOB);
begin
  if vTobEcr.GetInteger('E_NUMECHE') = 0 then Exit ;
  if vTobEcr.GetInteger('E_NUMLIGNE') <> NumLigne then Exit ;

  vTobEcr.ChangeParent( self,  0 ) ;

end;

procedure TMultiEche.CumuleTobEcr;

  function _GetCumulEche ( vStChamps : String ) : Double ;
  var i : Integer ;
  begin
    result := 0 ;
    for i := 0 to ( NbEche - 1 ) do
      result := result + Detail[ i ].GetDouble( vStChamps ) ;
  end ;

begin
    FTobEcr.PutValue( 'E_DEBIT',         _GetCumulEche('E_DEBIT' ) ) ;
    FTobEcr.PutValue( 'E_CREDIT',        _GetCumulEche('E_CREDIT') ) ;
    FTobEcr.PutValue( 'E_DEBITDEV',      _GetCumulEche('E_DEBITDEV') ) ;
    FTobEcr.PutValue( 'E_CREDITDEV',     _GetCumulEche('E_CREDITDEV') ) ;
    FTobEcr.PutValue( 'E_COUVERTURE',    _GetCumulEche('E_COUVERTURE') ) ;
    FTobEcr.PutValue( 'E_COUVERTUREDEV', _GetCumulEche('E_COUVERTUREDEV') ) ;
    FTobEcr.PutValue( 'E_ECHEENC1',      _GetCumulEche('E_ECHEENC1') ) ;
    FTobEcr.PutValue( 'E_ECHEENC2',      _GetCumulEche('E_ECHEENC2') ) ;
    FTobEcr.PutValue( 'E_ECHEENC3',      _GetCumulEche('E_ECHEENC3') ) ;
    FTobEcr.PutValue( 'E_ECHEENC4',      _GetCumulEche('E_ECHEENC4') ) ;
    FTobEcr.PutValue( 'E_ECHEDEBIT',     _GetCumulEche('E_ECHEDEBIT') ) ;
end;

function TMultiEche.GetEche(vNumEche: Integer) : TOB ;
begin
  result := nil ;
  if vNumEche > Detail.count
    then Exit ;
  result := Detail[ vNumEche - 1 ] ;
end;

procedure TMultiEche.Trier;
begin
  Detail.Sort('E_NUMECHE');
end;

function TMultiEche.EstLettre(vNumEche: Integer): Boolean;
begin
  result := False ;
  if IsOut( vNumEche) then Exit ;
  result := Trim( GetEche( vNumEche ).GetString('E_LETTRAGE') ) <> '' ;
end;

function TMultiEche.IsOut(vNumEche: Integer): Boolean;
begin
  result := vNumEche > Detail.Count ;
end;

function TMultiEche.EstModifiable(vNumEche: Integer): Boolean;
var i : Integer ;
begin
  result := True ;
  if vNumEche > 0
    then result := not EstLettre( vNumEche )
    else for i := 1 to Detail.Count do
           result := result and not EstLettre( i ) ;
end;


procedure TMultiEche.SynchroniseEcheances ;
var lTobEche : TOB ;
    lTobTemp : TOB ;
    i        : Integer ;
begin

  lTobTemp := TOB.Create('ECRITURE', nil, -1)  ;
  lTobTemp.InitValeurs ;

  for i := 1 to GetNbEche do
    begin
    lTobEche := GetEche( i ) ;

    // Mise en zone tampon des infos de l'échéance
    lTobTemp.Dupliquer ( lTobEche, False, True ) ;

    // Maj de l'échéance avec les données de l'écriture "globale"
    lTobEche.Dupliquer ( TobEcr, False, True ) ;

    // Récupération des infos spécifique à l'échéance :
    lTobEche.PutValue('E_NUMECHE',         i    ) ;
    PutTobEche( i, lTobTemp ) ;

    end ;

  FreeAndNil( lTobTemp ) ;

end;


function TMultiEche.GetModeRegl : T_MODEREGL;
var lTobEche : TOB ;
    i        : Integer ;
begin

  FillChar( result, Sizeof( result ), #0 ) ;

  if not EstModifiable
    then result.Action := taConsult
    else result.Action := taModif ;

  if InfoEcr.Compte.IsCollectif then
    begin
    result.Aux          := TobEcr.GetValue('E_AUXILIAIRE') ;
    result.ModeInitial  := InfoEcr.Aux_GetValue('T_MODEREGLE') ;
    end
  else
    begin
    result.Aux          := TobEcr.GetValue('E_GENERAL') ;
    result.ModeInitial  := InfoEcr.Compte_GetValue('G_MODEREGLE') ;
    end ;

  result.TotalAPayerP := TobEcr.GetValue('E_DEBIT')    + TobEcr.GetValue('E_CREDIT') ;
  result.TotalAPayerD := GetMontantEcr ;

  result.CodeDevise   := Devise.Code ;
  result.Symbole      := Devise.Symbole ;
  result.Quotite      := Devise.Quotite ;
  result.TauxDevise   := Devise.Taux ;
  result.Decimale     := Devise.Decimale ;

  result.DateFact     := GetDateFact ;
  result.DateBL       := GetDateBL ;
  result.DateFactExt  := GetDateFactExt ;

  result.JourPaiement1 := 0 ;
  result.JourPaiement2 := 0 ;

  result.NbEche       := NbEche ;
  result.ModeFinal    := ModeFinal ;

  // Détails
  for i := 1 to Detail.count do
    begin

    lTobEche := Detail[ i - 1 ] ;

    result.TabEche[i].MontantD       := GetMontantEche( i ) ;
    result.TabEche[i].MontantP       := lTobEche.GetValue('E_CREDIT')    + lTobEche.GetValue('E_DEBIT') ;
    result.TabEche[i].ReadOnly       := Trim( lTobEche.GetValue('E_LETTRAGE') ) <> '' ;
    result.TabEche[i].CodeLettre     := lTobEche.GetValue('E_LETTRAGE') ;
    result.TabEche[i].ModePaie       := lTobEche.GetValue('E_MODEPAIE') ;
    result.TabEche[i].DateEche       := lTobEche.GetValue('E_DATEECHEANCE') ;
    result.TabEche[i].DateValeur     := lTobEche.GetValue('E_DATEVALEUR') ;
    result.TabEche[i].DateRelance    := lTobEche.GetValue('E_DATERELANCE') ;
    result.TabEche[i].NiveauRelance  := lTobEche.GetValue('E_NIVEAURELANCE') ;
    result.TabEche[i].Couverture     := lTobEche.GetValue('E_COUVERTURE') ;
    result.TabEche[i].CouvertureDev  := lTobEche.GetValue('E_COUVERTUREDEV') ;
    result.TabEche[i].LettrageDev    := lTobEche.GetValue('E_LETTRAGEDEV') ;
    result.TabEche[i].DatePaquetMax  := lTobEche.GetValue('E_DATEPAQUETMAX') ;
    result.TabEche[i].DatePaquetMin  := lTobEche.GetValue('E_DATEPAQUETMIN') ;
    result.TabEche[i].EtatLettrage   := lTobEche.GetValue('E_ETATLETTRAGE') ;

    // Calcul du pourcentage...
    if result.TotalAPayerP <> 0 then
      result.TabEche[i].Pourc        := Arrondi( 100.0 * result.TabEche[i].MontantP / result.TotalAPayerP , V_PGI.OkDecP ) ;

    // #TVAENC
    if EstTvaLoc then
      begin
      result.TabEche[i].TAV[1] := lTobEche.GetValue( 'E_ECHEENC1' ) ;
      result.TabEche[i].TAV[2] := lTobEche.GetValue( 'E_ECHEENC2' ) ;
      result.TabEche[i].TAV[3] := lTobEche.GetValue( 'E_ECHEENC3' ) ;
      result.TabEche[i].TAV[4] := lTobEche.GetValue( 'E_ECHEENC4' ) ;
      result.TabEche[i].TAV[5] := lTobEche.GetValue( 'E_ECHEDEBIT' ) ;
      end ;

    end ; // fin For

end;

procedure TMultiEche.SetModeRegl( vModeRegl: T_MODEREGL );
var lTobEche : TOB ;
    lNumEche : Integer ;
begin

  // Suppression des anciennes échéances
  if NbEche > 0 then
    ClearDetail ;

  // Mode final
  ModeFinal := vModeRegl.ModeFinal ;

  // Parcours des écheances
  for lNumEche := 1 to vModeRegl.NbEche do
    begin
    lTobEche := GetEche( lNumEche ) ;
    if lTobEche = nil then
      lTobEche := NewEche( lNumEche ) ;

    if EstDebit then
      begin
      lTobEche.PutValue('E_DEBITDEV',     vModeRegl.TabEche[lNumEche].MontantD ) ;
      lTobEche.PutValue('E_DEBIT',        vModeRegl.TabEche[lNumEche].MontantP ) ;
      lTobEche.PutValue('E_CREDITDEV',    0 ) ;
      lTobEche.PutValue('E_CREDIT',       0 ) ;
      end
    else
      begin
      lTobEche.PutValue('E_DEBITDEV',  0 ) ;
      lTobEche.PutValue('E_DEBIT',     0 ) ;
      lTobEche.PutValue('E_CREDITDEV', vModeRegl.TabEche[lNumEche].MontantD ) ;
      lTobEche.PutValue('E_CREDIT',    vModeRegl.TabEche[lNumEche].MontantP ) ;
      end ;


    lTobEche.PutValue( 'E_MODEPAIE',        vModeRegl.TabEche[lNumEche].ModePaie      ) ;
    lTobEche.PutValue( 'E_DATEECHEANCE',    vModeRegl.TabEche[lNumEche].DateEche      ) ;
    if vModeRegl.TabEche[lNumEche].DateValeur > iDate1900
      then lTobEche.PutValue( 'E_DATEVALEUR',    vModeRegl.TabEche[lNumEche].DateValeur ) ;
    lTobEche.PutValue( 'E_ORIGINEPAIEMENT', vModeRegl.TabEche[lNumEche].DateEche      ) ;

    // Zone lettrage + relance non utilisée normalement en saisie
    lTobEche.PutValue( 'E_DATERELANCE',   vModeRegl.TabEche[lNumEche].DateRelance   ) ;
    lTobEche.PutValue( 'E_NIVEAURELANCE', vModeRegl.TabEche[lNumEche].NiveauRelance ) ;

    lTobEche.PutValue( 'E_LETTRAGE',      vModeRegl.TabEche[lNumEche].CodeLettre    ) ;
    lTobEche.PutValue( 'E_COUVERTURE',    vModeRegl.TabEche[lNumEche].Couverture    ) ;
    lTobEche.PutValue( 'E_COUVERTUREDEV', vModeRegl.TabEche[lNumEche].CouvertureDev ) ;
    lTobEche.PutValue( 'E_LETTRAGEDEV',   vModeRegl.TabEche[lNumEche].LettrageDev   ) ;
    lTobEche.PutValue( 'E_DATEPAQUETMAX', vModeRegl.TabEche[lNumEche].DatePaquetMax ) ;
    lTobEche.PutValue( 'E_DATEPAQUETMIN', vModeRegl.TabEche[lNumEche].DatePaquetMin ) ;
    lTobEche.PutValue( 'E_ETATLETTRAGE',  vModeRegl.TabEche[lNumEche].EtatLettrage  ) ;

    // Gestion de la tva
    if EstTvaLoc then
      begin
      lTobEche.PutValue( 'E_ECHEENC1' ,   vModeRegl.TabEche[lNumEche].TAV[1] ) ;
      lTobEche.PutValue( 'E_ECHEENC2' ,   vModeRegl.TabEche[lNumEche].TAV[2] ) ;
      lTobEche.PutValue( 'E_ECHEENC3' ,   vModeRegl.TabEche[lNumEche].TAV[3] ) ;
      lTobEche.PutValue( 'E_ECHEENC4' ,   vModeRegl.TabEche[lNumEche].TAV[4] ) ;
      lTobEche.PutValue( 'E_ECHEDEBIT',   vModeRegl.TabEche[lNumEche].TAV[5] ) ;
      end ;
    end ;

end;

function TMultiEche.GetDateFactExt: TDateTime;
begin
  result := TobEcr.GetDateTime('E_DATEREFEXTERNE') ;
  if result <= IDate1900 then
    result := TobEcr.GetDateTime('E_DATECOMPTABLE') ;
end;


function TMultiEche.GetMontantEche(vNumEche: Integer): Double;
var lTobEche : TOB ;
begin
  result := 0 ;
  lTobEche := GetEche( vNumEche ) ;
  if lTobEche = nil then Exit ;
  result := lTobEche.GetValue('E_CREDITDEV') + lTobEche.GetValue('E_DEBITDEV') ;
end;

function TMultiEche.EstTvaLoc: Boolean;
begin
  result := ( ( TobEcr.GetString('E_QUALIFPIECE') = 'N' ) or
              ( TobEcr.GetString('E_QUALIFPIECE') = 'S' ) or
              ( TobEcr.GetString('E_QUALIFPIECE') = 'R' ) )
            and ( VH^.OuiTvaEnc ) ;
end;

procedure TMultiEche.UpdateEche(vNumEche: Integer; vModePaie: String;  vMontant: Double; vDateEche: TDateTime);
var lTobEche : TOB ;
begin

  lTobEche := GetEche( vNumEche ) ;
  if lTobEche = nil then
    lTobEche := NewEche( vNumEche ) ;

  // MAJ mode de paiement
  lTobEche.PutValue( 'E_MODEPAIE',        vModePaie      ) ;
  lTobEche.PutValue( 'E_CODEACCEPT',      MPTOACC ( vModePaie ) ) ;

  // MAJ montant
  if EstDebit
    then CSetMontants( lTobEche,   vMontant,    0,           Devise,   True )
    else CSetMontants( lTobEche,   0,           vMontant,    Devise,   True ) ;

  // MAJ date d'échéance
  lTobEche.PutValue( 'E_DATEECHEANCE',    vDateEche  ) ;
  lTobEche.PutValue( 'E_ORIGINEPAIEMENT', vDateEche  ) ;

  // MAJ de l'écriture
  if vNumEche = 1 then
    SynchroniseEcr ;

  FBoModif := True ;

end;

function TMultiEche.GereArrondi: Boolean;
var lMontant  : Double ;
begin
  result := False ;
  if NbEche < 1 then Exit ;

  // Si le total des échéances n'est pas égale au montant de l'ecriture, on impacte la différence sur la dernière ligne
  if GetTotalEche <> GetMontantEcr then
    begin
    // Recalcul du montant de la ligne
    lMontant :=  GetMontantEche( NbEche - 1 ) + ( GetMontantEcr - GetTotalEche ) ;
    // MAJ du montant
    if EstDebit
      then CSetMontants( GetEche(NbEche-1),   lMontant,    0,           Devise,   True )
      else CSetMontants( GetEche(NbEche-1),   0,           lMontant,    Devise,   True ) ;
    // Indicateur de modification
    result := True ;
    end;

end;

procedure TMultiEche.GerePoucentage;
var lNumEche  : Integer ;
    lTobEche  : TOB ;
    lPourcent : Double ;
begin
  if GetMontantEcr <> 0 then
    for lNumEche := 1 to NbEche do
      begin
      lTobEche := GetEche( lNumEche ) ;
      lPourcent := Arrondi( 100.0 * GetMontantEche( lNumEche)  / GetMontantEcr, V_PGI.OkDecP );
      if lTobEche.GetNumChamp('POURCENTAGE') > 0
        then lTobEche.PutValue('POURCENTAGE', lPourcent )
        else lTobEche.AddChampSupValeur('POURCENTAGE', lPourcent ) ;
      end ;
end;

function TMultiEche.GetDevise: RDevise;
begin
  result := InfoEcr.Devise.Dev ;
end;

function TMultiEche.GetTotalEche: Double;
var lNumEche : Integer ;
begin
  result := 0 ;
  for lNumEche := 1 to NbEche do
    result := result + GetMontantEche( lNumEche ) ;
end;

procedure TMultiEche.Renumerote;
var lTobEche : TOB ;
    i        : Integer ;
begin

  for i := 1 to GetNbEche do
    begin
    lTobEche := GetEche( i ) ;
    // Récup nu numéro de lignes
    lTobEche.PutValue('E_NUMLIGNE',        TobEcr.GetInteger('E_NUMLIGNE')   ) ;
    // Récupération des infos spécifique à l'échéance :
    lTobEche.PutValue('E_NUMECHE',         i   ) ;
    end ;

end;

procedure TMultiEche.PutValueAll(vChamps: String; vValeur: Variant);
var i : Integer ;
begin
  for i := 0 to ( Detail.Count - 1 ) do
    Detail[ i ].PutValue( vChamps, vValeur ) ;
end;

function TMultiEche.GetDateBL: TDateTime;
begin
  result := TobEcr.GetDateTime('E_DATECOMPTABLE') ;
end;


function TMultiEche.EstOkMultiEche: Boolean;
begin
  Result := not (    ( InfoEcr.Compte_GetValue('G_NATUREGENE') = 'BQE' )
                  or ( InfoEcr.Compte_GetValue('G_NATUREGENE') = 'CAI' )
                  or   InfoEcr.Compte.IsVentilable
                  or ( InfoEcr.Journal.GetValue('J_EFFET') = 'X' )
                  or ( InfoEcr.Journal.GetValue('J_NATUREJAL') = 'BQE' )
                  or ( InfoEcr.Journal.GetValue('J_NATUREJAL') = 'CAI' )
                ) ;
end;

procedure TMultiEche.ChargeInfoEcr;
begin
  if not Assigned(FTobEcr) then Exit ;
  InfoEcr.Load( FTobEcr.GetString('G_GENERAL'), FTobEcr.GetString('E_AUXILIAIRE'), FTobEcr.GetString('E_JOURNAL') ) ;
end;

procedure TMultiEche.SetTvaEnc( vTvaEnc1, vTvaEnc2, vTvaEnc3, vTvaEnc4, vTvaDebit : Double ) ;
var lTobEche : TOB;
    lNumEche : Integer ;
    lTotal   : Double ;
    lCoef     : Double ;
    lMontant  : Double ;
begin

  if ModifTva then Exit ;

  lTotal := FTobEcr.GetValue('E_DEBIT') + FTobEcr.GetValue('E_CREDIT') ;
  if lTotal = 0 then Exit ;

  for lNumEche := 1 to NbEche do
    begin

    lTobEche  := GetEche( lNumEche ) ;

    // Calcul du coefficient de répartition sur l'échéance courante
    lCoef := (lTobEche.GetValue('E_DEBIT') + lTobEche.GetValue('E_CREDIT')) / lTotal ;

    // ECHE ENC 1
    lMontant := Arrondi( vTvaEnc1 * lCoef , V_PGI.OkDecV ) ;
    lTobEche.PutValue('E_ECHEENC1' ,     lMontant ) ;
    // ECHE ENC 2
    lMontant := Arrondi( vTvaEnc2 * lCoef , V_PGI.OkDecV ) ;
    lTobEche.PutValue('E_ECHEENC2' ,     lMontant ) ;
    // ECHE ENC 3
    lMontant := Arrondi( vTvaEnc3 * lCoef , V_PGI.OkDecV ) ;
    lTobEche.PutValue('E_ECHEENC3' ,     lMontant ) ;
    // ECHE ENC 4
    lMontant := Arrondi( vTvaEnc4 * lCoef , V_PGI.OkDecV ) ;
    lTobEche.PutValue('E_ECHEENC4' ,     lMontant ) ;
    // ECHE DEBIT
    lMontant := Arrondi( vTvaDebit * lCoef , V_PGI.OkDecV ) ;
    lTobEche.PutValue('E_ECHEDEBIT' ,     lMontant ) ;

    lTobEche.PutValue('E_EMETTEURTVA',   'X' ) ;

    end ;

  // MAJ de l'écriture de référence
  FTobEcr.PutValue( 'E_EMETTEURTVA',   'X' ) ;
  CumuleTobEcr ;


end;

procedure TMultiEche.InitTvaEnc;
var lTobEche : TOB;
    lNumEche : Integer ;
begin

  if ModifTva then Exit ;

  for lNumEche := 1 to NbEche do
    begin

    lTobEche  := GetEche( lNumEche ) ;

    // Initialisation des zones concernées
    lTobEche.PutValue( 'E_ECHEENC1',    0 ) ;
    lTobEche.PutValue( 'E_ECHEENC2',    0 ) ;
    lTobEche.PutValue( 'E_ECHEENC3',    0 ) ;
    lTobEche.PutValue( 'E_ECHEENC4',    0 ) ;
    lTobEche.PutValue( 'E_ECHEDEBIT',   0 ) ;
    lTobEche.PutValue( 'E_EMETTEURTVA', '-' ) ;

    end ;

end;

procedure TMultiEche.SynchroniseEcr;
begin
  if NbEche < 1 then Exit ;

  FTobEcr.PutValue( 'E_MODEPAIE',        Detail[0].GetValue('E_MODEPAIE') ) ;
  FTobEcr.PutValue( 'E_DATEECHEANCE',    Detail[0].GetValue('E_DATEECHEANCE') ) ;
  FTobEcr.PutValue( 'E_ORIGINEPAIEMENT', Detail[0].GetValue('E_ORIGINEPAIEMENT') ) ;
  FTobEcr.PutValue( 'E_CODEACCEPT',      Detail[0].GetValue('E_CODEACCEPT') ) ;

end;

procedure TMultiEche.CopieEcheances( vSource : TMultiEche );
var i : Integer ;
begin
  FStModeFinal := vSource.ModeFinal ;
  FBoModifTVA  := vSource.ModifTVA ;
  FBoModif     := vSource.ModifEnCours ;

  ChargeMR ;

  ClearDetail ;
  for i := 1 to vSource.NbEche do
    begin
    NewEche ;
    PutTobEche( i, vSource.GetEche( i )  ) ;
    end ;

end;

procedure TMultiEche.PutTobEche ( vNumEche : Integer ; vTobSource : TOB ) ;
var lTobEche : TOB ;
begin

  lTobEche := GetEche( vNumEche ) ;
  if lTobEche = nil then Exit ;

  lTobEche.PutValue('E_ECHE',            'X'  ) ;

  lTobEche.PutValue('E_NIVEAURELANCE',   vTobSource.GetValue('E_NIVEAURELANCE') ) ;
  lTobEche.PutValue('E_MODEPAIE',        vTobSource.GetValue('E_MODEPAIE')      ) ;
  lTobEche.PutValue('E_DATEECHEANCE',    vTobSource.GetValue('E_DATEECHEANCE')  ) ;
  lTobEche.PutValue('E_DATEVALEUR',      vTobSource.GetValue('E_DATEVALEUR')    ) ;
  lTobEche.PutValue('E_COUVERTURE',      vTobSource.GetValue('E_COUVERTURE')    ) ;
  lTobEche.PutValue('E_COUVERTUREDEV',   vTobSource.GetValue('E_COUVERTUREDEV') ) ;
  lTobEche.PutValue('E_ETATLETTRAGE',    vTobSource.GetValue('E_ETATLETTRAGE')  ) ;
  lTobEche.PutValue('E_LETTRAGE',        vTobSource.GetValue('E_LETTRAGE')      ) ;
  lTobEche.PutValue('E_LETTRAGEDEV',     vTobSource.GetValue('E_LETTRAGEDEV')   ) ;
  lTobEche.PutValue('E_DATEPAQUETMAX',   vTobSource.GetValue('E_DATEPAQUETMAX') ) ;
  lTobEche.PutValue('E_DATEPAQUETMIN',   vTobSource.GetValue('E_DATEPAQUETMIN') ) ;
  lTobEche.PutValue('E_DATERELANCE',     vTobSource.GetValue('E_DATERELANCE')   ) ;
  lTobEche.PutValue('E_DEBITDEV',        vTobSource.GetValue('E_DEBITDEV')      ) ;
  lTobEche.PutValue('E_DEBIT',           vTobSource.GetValue('E_DEBIT')         ) ;
  lTobEche.PutValue('E_CREDITDEV',       vTobSource.GetValue('E_CREDITDEV')     ) ;
  lTobEche.PutValue('E_CREDIT',          vTobSource.GetValue('E_CREDIT')        ) ;
  lTobEche.PutValue('E_ORIGINEPAIEMENT', vTobSource.GetValue('E_DATEECHEANCE')  ) ;
  lTobEche.PutValue('E_ENCAISSEMENT',    vTobSource.GetValue('E_ENCAISSEMENT')  ) ;
  lTobEche.PutValue('E_CODEACCEPT',      MPTOACC ( vTobSource.GetValue('E_MODEPAIE') ) ) ;
  // TVA
  lTobEche.PutValue( 'E_ECHEENC1',       vTobSource.GetValue('E_ECHEENC1') ) ;
  lTobEche.PutValue( 'E_ECHEENC2',       vTobSource.GetValue('E_ECHEENC2') ) ;
  lTobEche.PutValue( 'E_ECHEENC3',       vTobSource.GetValue('E_ECHEENC3') ) ;
  lTobEche.PutValue( 'E_ECHEENC4',       vTobSource.GetValue('E_ECHEENC4') ) ;
  lTobEche.PutValue( 'E_ECHEDEBIT',      vTobSource.GetValue('E_ECHEDEBIT') ) ;
  lTobEche.PutValue( 'E_EMETTEURTVA',    vTobSource.GetValue('E_EMETTEURTVA') ) ;
  // ??
  lTobEche.PutValue('E_NUMTRAITECHQ',    vTobSource.GetValue('E_NUMTRAITECHQ')  ) ;

  // Gestion Cut-Off
  if vTobSource.GetNumChamp('PCOMPL') > 0 then
    begin
    if lTobEche.GetNumChamp('PCOMPL') <= 0 then
      lTobEche.AddChampSup('PCOMPL', False ) ;
    lTobEche.PutValue( 'PCOMPL', lTobEche.GetValue('PCOMPL') ) ;
    end ;

end;

end.
