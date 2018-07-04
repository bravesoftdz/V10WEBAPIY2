{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 19/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPLOIVENTILMUL ()
Mots clefs ... : TOF;CPLOIVENTILMUL
*****************************************************************}
Unit CPLOIVENTILMUL_TOF ;

Interface

Uses StdCtrls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
//     mul,
{$else}
//     eMul,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HQry,       // THQuery
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF ;

Type
  TOF_CPLOIVENTILMUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure OnAfterShow;
  private

    // Composants
    Q                 : THQuery ;
    CLV_AXE           : THValComboBox ;

    // Indicateurs
    FStType           : string;                // Stocke le type de répartition concernée
    FBoChargEnCours   : boolean ;

    // Initialisation de la fiche
    procedure AnalyseParametre(vStArg: string);
    procedure InitComposants;
    procedure InitEvenements;
    procedure InitEcran;

    // Evènement de la fiche
    procedure FListeDblClick   ( Sender : TObject );
    procedure BInsertClick     ( Sender : TObject );
    procedure BChercheClick    ( Sender : TObject );
    procedure BSupprimerClick  ( Sender : TObject );

    // Evènement des critères
    procedure AxeChanged       ( Sender : TObject );

    procedure SetDataType      ( vStAxe : string ) ;
    function  SelectionOk      : Boolean ;

  end ;

function CPLanceFiche_LoiVentilMul( vStRange, vStLequel, vStArgs : string ) : string;

Implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  windows,
  {$IFDEF VER150} Variants, {$ENDIF}    // VarToStr
  {$IFDEF EAGLCLIENT}
     MainEagl,          // AglLanceFiche
     eMul,
  {$ELSE}
     mul,
     FE_Main,           // AglLanceFiche
  {$ENDIF}
  HTB97,                // TToolBarButton97
  Ent1,                 // GetInfoCpta
  CPLOIVENTIL_TOF,      // CPLOIVENTIL_TOF
  Controls;             // mrYes



// =============================================================================
function CPLanceFiche_LoiVentilMul( vStRange, vStLequel, vStArgs : string ) : string;
begin
  result := AGLLanceFiche('CP', 'CPLOIVENTILMUL', vStRange, vStLequel, vStArgs ) ;
end ;
// =============================================================================


procedure TOF_CPLOIVENTILMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPLOIVENTILMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPLOIVENTILMUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPLOIVENTILMUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPLOIVENTILMUL.OnArgument (S : String ) ;
begin
  FBoChargEnCours := true;

  Inherited ;

  AnalyseParametre(S);
  InitComposants;
  InitEvenements;
  InitEcran;

end ;

procedure TOF_CPLOIVENTILMUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPLOIVENTILMUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPLOIVENTILMUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPLOIVENTILMUL.AnalyseParametre(vStArg: string);
var
  lStChaine: string;
begin

  while vStArg <> '' do
    begin
    lStChaine := ReadTokenSt(vStArg);

    // Type de répartition concernée
    if Pos('CLV_LOITYPE=', Uppercase(lStChaine)) > 0 then
      FStType := Copy(lStChaine, 13, length(lStChaine))
    ;

    end;

  if FStType='' then
    FStType := 'SEC' ;

end;

procedure TOF_CPLOIVENTILMUL.InitComposants;
begin
  // Composants
  Q       := THQuery(GetControl('Q', True));
  CLV_AXE := THValComboBox( GetControl('CLV_AXE', True) ) ;
end;

procedure TOF_CPLOIVENTILMUL.InitEcran;
begin

  // Critère invisible
  SetControlVisible('CLV_LOITYPE', False ) ;
  SetControlText('CLV_LOITYPE', FStType ) ;
  SetControlVisible('BSUPPRIMER', True ) ;

  if FStType = 'LOI' then
    begin
    // Liste décla : loi de répartition
    TFMul(Ecran).SetDBListe('CPLOIMODELE');
    TTabSheet( GetControl('PCOMPLEMENT', TRue) ).TabVisible := False ;
    Ecran.caption := 'Lois de répartition' ;
    TCheckBox(GetControl('CLV_STRICTE')).State := cbGrayed ;
    end
  else
    begin
    // Liste décla : répartition secondaire
    TFMul(Ecran).SetDBListe('CPLOIVENTIL');
    TTabSheet( GetControl('PCOMPLEMENT', TRue) ).TabVisible := True ;
    Ecran.caption := 'Répartitions secondaires' ;
    SetControlVisible('CLV_STRICTE', False ) ;
    SetControlVisible('TCLV_STRICTE', False ) ;
    TCheckBox(GetControl('CLV_STRICTE')).State := cbGrayed ;

    // valeur par défaut
    THEdit( GetControl('CLV_DATEGENER')).Operateur := Egal ;
    THEdit( GetControl('CLV_DATETRAITEMENT')).Operateur := Egal ;
    SetControlText('CLV_DATEGENER',      stDate1900 ) ;
    SetControlText('CLV_DATETRAITEMENT', stDate1900 ) ;
    end ;

  if CLV_AXE.ItemIndex < 0 then
    CLV_AXE.ItemIndex := 0 ;
  AxeChanged(nil);

  UpdateCaption( Ecran ) ;

end;

procedure TOF_CPLOIVENTILMUL.InitEvenements;
begin

  TFMul(Ecran).OnAfterFormShow := OnAfterShow;

  // Double click sur la liste
  TFMul(Ecran).FListe.OnDblClick := FListeDblClick;

  // Bouton de création
  TToolBarButton97(GetControl('BINSERT', True)).OnClick    := BInsertClick;
  TToolBarButton97(GetControl('BCHERCHE', True)).OnClick   := BChercheClick;
  TToolBarButton97(GetControl('BSUPPRIMER', True)).OnClick := BSupprimerClick;

  // Critères
  CLV_AXE.OnChange                                         := AxeChanged ;

end;

procedure TOF_CPLOIVENTILMUL.BChercheClick(Sender: TObject);
begin

  TFMul(Ecran).BChercheClick(Sender) ;
  Exit ;
  // Effectué la recherche uniquement si
  if not FBoChargEnCours and ( GetControlText('CLV_SECTION') <> '' ) then
    begin
    TFMul(Ecran).BChercheClick(Sender)
    end
{$IFDEF EAGLCLIENT}
  ;
{$ELSE EAGLCLIENT}
  else
    if FBoChargEnCours then
      if TFMUL(Ecran).HMTrad.ActiveResize and TFMUL(Ecran).HMTrad.ResizeDBGrid and V_PGI.Outlook then
        TFMUL(Ecran).HMTrad.ResizeDBGridColumns(TFMUL(Ecran).FListe);
{$ENDIF EAGLCLIENT}


end;

procedure TOF_CPLOIVENTILMUL.BInsertClick(Sender: TObject);
var lStArg : string ;
    lStAxe : string ;
    lStSec : string ;
    lStGen : string ;
begin

  // Test axe sélectionné
  lStAxe := GetControlText('CLV_AXE') ;
  if lStAxe = '' then
    begin
    PgiInfo( TraduireMemoire('Vous devez sélectionner un axe.'), Ecran.Caption ) ;
    SetFocusControl('CLV_AXE') ;
    Exit ;
    end ;

  // Test journal d'ODA sur axe sélectionné
  if not ExisteSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_FERME="-" AND J_NATUREJAL="ODA" AND J_AXE="' + lStAxe + '"') then
    begin
    PgiInfo( TraduireMemoire('Il n''existe aucun journal d''ODA ouvert pour l''axe sélectionné.'), Ecran.Caption ) ;
    Exit ;
    end ;

  // Test section d'attente sur axe sélectionné
  lStSec := GetInfoCpta( AxeToFb( lStAxe )).Attente ;
  if Trim(lStSec) = '' then
    begin
    PgiInfo( TraduireMemoire('La section d''attente n''est pas paramétrée sur l''axe sélectionné.'), Ecran.Caption ) ;
    Exit ;
    end
  // test d'existence de la section paramétré
  else
    if not existeSQL( 'SELECT S_SECTION FROM SECTION WHERE S_SECTION="' + lStSec + '"') then
      begin
      PgiInfo( TraduireMemoire('La section d''attente paramétrée sur l''axe sélectionné n''existe pas.'), Ecran.Caption ) ;
      Exit ;
      end ;

  // Test compte géné d'attente sur axe sélectionné
  lStGen := GetInfoCpta( AxeToFb( lStAxe )).AxGenAttente ;
  if Trim(lStGen) = '' then
    begin
    PgiInfo( TraduireMemoire('Le compte général d''attente n''est pas paramétré sur l''axe sélectionné.'), Ecran.Caption ) ;
    Exit ;
    end
  // test d'existence du général paramétré
  else
    if not existeSQL( 'SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="' + lStGen + '"') then
      begin
      PgiInfo( TraduireMemoire('Le compte général d''attente paramétré sur l''axe sélectionné n''existe pas.'), Ecran.Caption ) ;
      Exit ;
      end ;

  // Arguments
  lStArg := 'ACTION=CREATION;CLV_LOITYPE=' + FStType + ';CLV_AXE=' + lStAxe ;
  lStSec := GetControlText('CLV_SECTION') ;
  if lStSec<>'' then
    begin
    if ExisteSQL('SELECT S_SECTION FROM SECTION WHERE S_AXE="' + lStAxe + '" '
                         + 'AND S_SECTION="' + lStSec + '" '
                         + 'AND S_INDIRECTE="X"') then
      lStArg := lStArg + ';CLV_SECTION=' + lStSec ;
    end ;

  // ouverture fiche
  CPLanceFiche_LoiVentil( '', '', lStArg) ;

  // relance la recherche
  BChercheClick(nil) ;
end;

procedure TOF_CPLOIVENTILMUL.FListeDblClick(Sender: TObject);
var lStArg : string ;
begin

  if not SelectionOk then
    Exit ;

  lStArg := 'ACTION=MODIFICATION;CLV_LOITYPE=' + FStType
             + ';CLV_AXE='    + GetField('CLV_AXE')
             + ';CLV_NUMERO=' + IntToStr(GetField('CLV_NUMERO')) ;

  CPLanceFiche_LoiVentil( '', '', lStArg) ;
  BChercheClick(nil) ;
end;


procedure TOF_CPLOIVENTILMUL.OnAfterShow;
begin
  FBoChargEnCours := False;
end;

procedure TOF_CPLOIVENTILMUL.BSupprimerClick(Sender: TObject);
begin

  if not SelectionOk then
    Exit ;

  // Confirmation
  if FStType = 'LOI' then
    begin
    if PGIAsk(TraduireMemoire('Confirmez-vous la suppression de la loi de répartition sélectionnée (') + GetField('CLV_LIBELLE') + ') ?',
              Ecran.Caption) <> mrYes then Exit ;
    end
  else
    begin
    if PGIAsk(TraduireMemoire('Confirmez-vous la suppression de l''historique de la répartition secondaire sélectionnée (') + GetField('CLV_LIBELLE') + ') ?',
              Ecran.Caption) <> mrYes then Exit ;
    end ;

  // Supppression
  if executeSql('DELETE FROM CLOIVENTIL WHERE CLV_LOITYPE="' + FStType + '" AND CLV_NUMERO=' + VarToStr(GetField('CLV_NUMERO')) ) <> 1
    then PGIInfo(TRaduireMemoire('L''enregistrement n''a pas pu être supprimé.'), Ecran.Caption)
    else executeSql('DELETE FROM VENTIL WHERE V_NATURE="CLV" AND V_COMPTE="' + VarToStr(GetField('CLV_VENTIL')) + '"' ) ;

  // On relnce la recherche
  TFMul(Ecran).BChercheClick(nil);

end;

procedure TOF_CPLOIVENTILMUL.AxeChanged(Sender: TObject);
var lStAxe : string ;
begin
  lStAxe := GetControlText('CLV_AXE') ;

  // MAJ des DataType section et journal ODA...
  SetDataType( lStAxe ) ;

end;

procedure TOF_CPLOIVENTILMUL.SetDataType(vStAxe: string);
begin
  if length(vStAxe)<>2 then Exit ;
  SetControlProperty( 'CLV_SECTION',    'PLUS',     ' S_AXE="' + vStAxe + '" AND S_FERME="-" AND S_INDIRECTE="X" ' ) ;
  SetControlProperty( 'CLV_JOURNALODA', 'PLUS',     ' J_NATUREJAL="ODA" AND J_AXE="' + vStAxe + '" AND J_FERME="-" ' ) ;
end;

function TOF_CPLOIVENTILMUL.SelectionOk: Boolean;
begin

  result := False ;

  {$IFDEF EAGLCLIENT}
  if (Q.TQ.Detail.Count = 0) then
  {$ELSE}
  if (Q.Eof and Q.Bof) then
  {$ENDIF}
    begin
    PGIInfo( TraduireMemoire('Le résultat de la recherche est vide. Veuillez relancer une recherche.'), Ecran.Caption);
    exit;
    end;

  result := True ;

end;

Initialization
  registerclasses ( [ TOF_CPLOIVENTILMUL ] ) ;
end.
