unit ULibEcrSimplifiee;

interface

uses
  UTof, HCtrls, HTB97, StdCtrls, ComCtrls, HPanel, Classes, uTob;

const
  QUALIFSIMPLIFIE = 'SMP';

  {Ordre de champs de la TOB de InitStructureSimplifie}
  CHP_JOURNAL  = 1000;
  CHP_DATECPT  = 1001;
  CHP_GENERAL  = 1002;
  CHP_TIERS    = 1003;
  CHP_NATPIECE = 1004;
  CHP_DEVISE   = 1005;
  CHP_ETAB     = 1006;
  CHP_QUALIFP  = 1007;
  CHP_LIBELLE  = 1008;
  CHP_REFINT   = 1009;
  CHP_CREDITD  = 1010;
  CHP_DEBITD   = 1011;
  CHP_DATEVAL  = 1012;
  CHP_MODEPAIE = 1013;
  CHP_RIB      = 1014;


type  
  {Ancêtre des classes UTOFECRLET.TOF_ECRLET et CPSAISIEPIECE_TOF.TOF_CPSAISIEPIECE.
   Ces deux classes doivent reposer sur la même ergonomie => toute modification sur une
   fiche doit être reporté sur l'autre.
   On ne gère dans l'ancêtre que les composants, car si le but des deux fiches est de créer
   simplement et rapidement une pièce, le mécanisme est radicalement différents :
   - CECRSIMPL part d'une pièce de contrepartie et est utilisée dans le lettrage et la consultation des écritures
   - CPECRITURESIMPLE se veut une saisie beaucoup plus générique et repose sur le TPieceCompta : dans l'absolu
     il est possible de créer une pièce avec simplement un montant, une devise et un général}
  TofAncetreEcr = class(TOF)
    procedure OnArgument(S : string); override;
  protected
    {contrôles de la fiche}
    EdtDateComptable : THCritMaskEdit;
    EdtNaturePiece   : THCritMaskEdit;
    EdtEtablissement : THCritMaskEdit;
    EdtAuxiliaire    : THCritMaskEdit;
    lblAuxiliaire    : TLabel;
    lblFolio         : TLabel;
    EdtFolio         : THValComboBox;
    EdtLibelle       : THCritMaskEdit;
    EdtTva           : THCritMaskEdit;
    lblTva           : THLabel;
    EdtGeneral       : THCritMaskEdit;
    EdtJournal       : THCritMaskEdit;
    PgcDetail        : TPageControl;
    BtnVoir          : TToolbarButton97;
    G                : THGrid;
    P1               : THPanel;
    P2               : THPanel;
    P3               : THPanel;

    FStNatGene : string;

    function  GetControlTOF : Boolean;
    function  AssignEvent   : Boolean; virtual;
    function  IsValidDateComptable(Parle : Boolean = False) : Boolean; virtual;
    procedure RemplitComboFolio;
    procedure EdtAuxiliaireElipsisClick(Sender : TObject); virtual;
    procedure EdtJournalElipsisClick   (Sender : TObject); virtual; abstract;
    procedure EdtDateComptableExit     (Sender : TObject); virtual; abstract;
    procedure BtnVoirClick             (Sender : TObject); virtual; abstract;
    procedure EdtFolioExit             (Sender : TObject); virtual; abstract;
    procedure EdtFolioKeyPress         (Sender : TObject; var Key : Char); virtual;
    procedure FormKeyDown              (Sender : TObject; var Key : Word; Shift : TShiftState); virtual;
  end;

{Ajout des champs dans la tob pour les écritures simplifiées, dans l'ordre des constantes CHP_}
procedure InitStructureSimplifie(var aTob : TOB);


implementation

uses
  ULibEcriture, SaisUtil, Windows, SysUtils, Vierge, HEnt1;

{---------------------------------------------------------------------------------------}
procedure TofAncetreEcr.OnArgument(S: string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  GetControlTOF;
  AssignEvent;
end;

{---------------------------------------------------------------------------------------}
function TofAncetreEcr.GetControlTOF : Boolean;
{---------------------------------------------------------------------------------------}
begin
  EdtDateComptable := THCritMaskEdit(GetControl('E_DATECOMPTABLE'));
  EdtNaturePiece   := THCritMaskEdit(GetControl('E_NATUREPIECE'));
  EdtEtablissement := THCritMaskEdit(GetControl('E_ETABLISSEMENT'));
  EdtLibelle       := THCritMaskEdit(GetControl('E_LIBELLE'));
  EdtTva           := THCritMaskEdit(GetControl('E_TVA'));
  EdtAuxiliaire    := THCritMaskEdit(GetControl('E_AUXILIAIRE'));
  EdtGeneral       := THCritMaskEdit(GetControl('E_GENERAL'));
  EdtJournal       := THCritMaskEdit(GetControl('E_JOURNAL'));
  EdtFolio         := THValComboBox(GetControl('E_FOLIO'));
  PgcDetail        := TPageControl(GetControl('FE__PAGECONTROLDETAIL'));
  BtnVoir          := TToolbarButton97(GetControl('BVOIR'));
  G                := THGrid (GetControl('G' , True));
  lblTva           := THlabel(GetControl('TE_TVA'));
  lblAuxiliaire    := THlabel(GetControl('TE_AUXILIAIRE'));
  lblFolio         := THlabel(GetControl('TE_FOLIO'));
  P1               := THPanel(GetControl('P1', True));
  P2               := THPanel(GetControl('P2', True));
  P3               := THPanel(GetControl('P3', True));

  Result := True;
end;

{---------------------------------------------------------------------------------------}
function TofAncetreEcr.AssignEvent : Boolean;
{---------------------------------------------------------------------------------------}
begin
  EdtAuxiliaire.OnElipsisClick  := EdtAuxiliaireElipsisClick;
  EdtFolio.OnKeyPress           := EdtFolioKeyPress ;
  Ecran.OnKeyDown               := FormKeyDown;

  EdtJournal.OnElipsisClick     := EdtJournalElipsisClick;
  EdtDateComptable.OnExit       := EdtDateComptableExit;
  BtnVoir.OnClick               := BtnVoirClick;
  EdtFolio.OnExit               := EdtFolioExit ;

  Result := True;
end;

{---------------------------------------------------------------------------------------}
function TofAncetreEcr.IsValidDateComptable(Parle : Boolean = False) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := (Length(Trim(EdtDateComptable.Text))) = 10;
end;

{---------------------------------------------------------------------------------------}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 20/07/2007
Modifié le ... :   /  /    
Description .. : - LG - 20/07/2007 - FB 20202 - rajout d'un test sur les 
Suite ........ : ecritrue valide
Mots clefs ... : 
*****************************************************************}
procedure TofAncetreEcr.RemplitComboFolio;
{---------------------------------------------------------------------------------------}
begin
  if IsValidDateComptable then begin
    CRempliComboFolio ( EdtFolio.Items, EdtFolio.Values , EdtJournal.Text ,QuelExo(EdtDateComptable.Text) , strToDate(EdtDateComptable.Text) , true ) ;
    EdtFolio.ItemIndex := 0 ;
    if EdtFolio.Text = '' then
      EdtFolio.Text := '1' ;
  end
  else if EdtDateComptable.CanFocus then
    EdtDateComptable.Setfocus;
end;

{---------------------------------------------------------------------------------------}
procedure TofAncetreEcr.EdtAuxiliaireElipsisClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if (csDestroying in Ecran.ComponentState) then Exit;

  CLookupListAux(EdtAuxiliaire, Copy(EdtGeneral.Text, 1, 1), EdtNaturePiece.Text, FStNatGene);
end;

{---------------------------------------------------------------------------------------}
procedure TofAncetreEcr.EdtFolioKeyPress(Sender : TObject; var Key : Char);
{---------------------------------------------------------------------------------------}
begin
  if (Ord(Key) >= 32) and (not (Key in ['0'..'9'])) then Key := #0;
end;

{---------------------------------------------------------------------------------------}
procedure TofAncetreEcr.FormKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin
  case Key of
    VK_F6  : begin
               BtnVoirClick(BtnVoir);
               Key := 0;
             end;
    VK_F10 : begin
               TFVierge(Ecran).BValider.Click;
               Key := 0;
             end;
  end;
end;

{Ajout des champs dans la tob pour les écritures simplifiées.
{---------------------------------------------------------------------------------------}
procedure InitStructureSimplifie(var aTob : TOB);
{---------------------------------------------------------------------------------------}
begin
  {Création des champs correspondant aux properties de l'objet CPECRITURESIMPLE_TOF.TObjEcritureSimplif
   Pour créer une nouvelle property à l'objet, il faut
   1/ Créer la property dans la definition de la classe TObjEcritureSimplif
   2/ Créer une nouvelle constante (CHP_XXX) dans l'interface de ULibPointage
   3/ Ajouter le champ en respectant bien l'ordre des champs supplémentaires de la TOB}
  aTob.AddChampSupValeur('JOURNAL'  , ''       ); { CHP_JOURNAL  = 1000; }
  aTob.AddChampSupValeur('DATECPT'  , iDate1900); { CHP_DATECPT  = 1001; }
  aTob.AddChampSupValeur('GENERAL'  , ''       ); { CHP_GENERAL  = 1002; }
  aTob.AddChampSupValeur('TIERS'    , ''       ); { CHP_TIERS    = 1003; }
  aTob.AddChampSupValeur('NATPIECE' , ''       ); { CHP_NATPIECE = 1004; }
  aTob.AddChampSupValeur('DEVISE'   , ''       ); { CHP_DEVISE   = 1005; }
  aTob.AddChampSupValeur('ETAB'     , ''       ); { CHP_ETAB     = 1006; }
  aTob.AddChampSupValeur('QUALIFP'  , ''       ); { CHP_QUALIFP  = 1007; }
  aTob.AddChampSupValeur('LIBELLE'  , ''       ); { CHP_LIBELLE  = 1008; }
  aTob.AddChampSupValeur('REFINT'   , ''       ); { CHP_REFINT   = 1009; }
  aTob.AddChampSupValeur('CREDITD'  , 0.0      ); { CHP_CREDITD  = 1010; }
  aTob.AddChampSupValeur('DEBITD'   , 0.0      ); { CHP_DEBITD   = 1011; }
  aTob.AddChampSupValeur('DATEVAL'  , iDate1900); { CHP_DATEVAL  = 1012; }
  aTob.AddChampSupValeur('MODEPAIE' , ''       ); { CHP_MODEPAIE = 1013; }
  aTob.AddChampSupValeur('RIB'      , ''       ); { CHP_RIB      = 1014; }

end;

end.
