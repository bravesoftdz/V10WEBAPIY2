{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 28/12/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPGENEECRTP ()
Mots clefs ... : TOF;CPGENEECRTP
*****************************************************************}
Unit CPGENEECRTP_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     MaineAGL,
     eMul,
{$ELSE}
     db,
     Hdb,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,
     Mul,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HPanel,
     HQry,
     HTB97,                 //SG6 FQ 14202  23/11/04
     uTOF,
     uTOB,                  // TOB
     SAISUTIL,		    // Pour RMVT
     ParamSoc,              // Pour le GetParamSoc
     UtilPGI,               // Pour les procedures de blocage
{$IFDEF CCSTD}
{$ELSE}
  {$IFDEF COMPTA}
     Saisie,		    // Pour Saisie eAGL
     SaisBor,               // LanceSaisieFolio
     CPSaisiePiece_Tof,     // saisie paramètrable
     CPOBJENGAGE , //fb 02/05/2006
  {$ENDIF}
{$ENDIF}
     Ent1,		              // Pour EstMonaieIn et GetPeriode
     Ed_Tools,              // Pour le videListe
     TiersPayeur,           // Pour les fonctions xxxTP
     HStatus,               // Pour la barre d'état
     Zcompte,
     SaisComm,
     dialogs,              // Pour les procédures de MAJ des comptes
     HSysMenu,
     Windows,              // VK_F10
     ULibExercice,         // CQuelExercice
     UTobDebug,
     DelVisuE;

procedure CPLanceFiche_CPGeneEcrTP(vStRange, vStLequel, vStArgs : string);

Type
  TOF_CPGENEECRTP = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
      FEcran    : TFMul ;

    ANouveau  : Boolean ;
    MsgBox    : THMsgBox;
    HMTrad: THSystemMenu;

    // Eléments interface
    E_AUXILIAIRE  : THEdit ;
    CATEGORIE        : THValComboBox ;
    E_MODEPAIE        : THValComboBox ;
    E_JOURNAL        : THValComboBox ;
    E_NATUREPIECE    : THValComboBox ;
    E_DEVISE         : THValComboBox ;
//  E_QUALIFPIECE    : THValComboBox ;
    E_EXERCICE       : THValComboBox ;
    E_ETABLISSEMENT  : THValComboBox ;
    E_DATECOMPTABLE  : THEdit ;
    E_DATECOMPTABLE_ : THEdit ;
    E_DATEECHEANCE   : THEdit ;
    E_DATEECHEANCE_  : THEdit ;
    XX_WHERE         : THEdit ;
    INVISIBLE        : TTabSheet ;
  {$IFDEF EAGLCLIENT}
    FListe : THGrid ;
  {$ELSE}
    FListe : THDBGrid ;
  {$ENDIF}
  // Evènements fiche
    procedure E_EXERCICEChange(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure OuvrirClick ;
    function  AddFille (T : TOB ; Q : TQuery) : TOB;
    procedure ValideClick ;
    procedure BOuvrirClick ( Sender : TObject ) ;
    procedure BValideClick ( Sender : TObject ) ;
    procedure BChercheClick(Sender: TObject);
// Autres procédures
    Procedure InitFormModification ;
    Function ExisteTP (Q : TQuery) : Boolean ;
    procedure InitMsgBox;
    procedure FormKeyDown( Sender : TObject; var Key: Word; Shift: TShiftState );
    procedure AuxiElipsisClick         ( Sender : TObject );
    public
    Action : TActionFiche ;
  end ;

Implementation

uses

{$IFDEF EAGLCLIENT}
{$ELSE}
{$ENDIF}  // Bouton chgt de taux actif en eAGL YMO 01/06
//  RepDevEur, // ChangeLeTauxDevise
  UlibEcriture,
  {$IFDEF eAGLCLIENT}
  MenuOLX
  {$ELSE}
  MenuOLG
  {$ENDIF eAGLCLIENT}
  , UTofMulParamGen; {13/04/07 YMO F5 sur Auxiliaire }

type ttCarAlpha = set of char;
var StopCherche : boolean;


procedure CPLanceFiche_CPGeneEcrTP(vStRange, vStLequel, vStArgs : string);
begin

    AGLLanceFiche ('CP','CPGENEECRTP', vStRange, vStLequel, vStArgs );

end;

procedure TOF_CPGENEECRTP.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPGENEECRTP.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPGENEECRTP.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPGENEECRTP.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPGENEECRTP.OnArgument (S : String ) ;
begin

 // Récup interface
  FEcran := TFMul(Ecran) ;
{$IFDEF EAGLCLIENT}
  FListe  := THGrid( GetControl('FListe',True) ) ;
{$ELSE}
  FListe  := THDBGrid( GetControl('FListe',True)) ;
{$ENDIF}

//  GetControl('E_VALIDE',True).Enabled:=False;

  E_AUXILIAIRE     := THEdit(GetControl('E_AUXILIAIRE', True)) ;
  E_MODEPAIE        := THValComboBox(GetControl('E_MODEPAIE', True)) ;
  CATEGORIE        := THValComboBox(GetControl('CATEGORIE', True)) ;
  E_JOURNAL        := THValComboBox(GetControl('E_JOURNAL', True)) ;
  E_NATUREPIECE    := THValComboBox(GetControl('E_NATUREPIECE', True)) ;
  E_DEVISE         := THValComboBox(GetControl('E_DEVISE', True)) ;
//  E_QUALIFPIECE    := THValComboBox(GetControl('E_QUALIFPIECE', True)) ;
  E_EXERCICE       := THValComboBox(GetControl('E_EXERCICE', True)) ;
  E_ETABLISSEMENT  := THValComboBox(GetControl('E_ETABLISSEMENT', True)) ;
  E_DATECOMPTABLE  := THEdit(GetControl('E_DATECOMPTABLE', True))  ;
  E_DATECOMPTABLE_ := THEdit(GetControl('E_DATECOMPTABLE_', True)) ;
  E_DATEECHEANCE   := THEdit(GetControl('E_DATEECHEANCE', True)) ;
  E_DATEECHEANCE_  := THEdit(GetControl('E_DATEECHEANCE_', True)) ;
  XX_WHERE         := THEdit(GetControl('XX_WHERE', True)) ;
  INVISIBLE        := TTabsheet(GetControl('INVISIBLE', True)) ;

  TFMul(Ecran).OnKeyDown  := FormKeyDown;

	// Réaffectation des évènements
  E_EXERCICE.OnChange  := E_EXERCICEChange ;
  E_EXERCICE.Vide      :=True;
  E_EXERCICE.VideString:='<<Tous>>';

  FListe.OnDblClick := FListeDblClick ;

  INVISIBLE.TabVisible:=False ;

  TButton(GetControl('BOuvrir', True)).OnClick 		  := BOuvrirClick ;
  TButton(GetControl('BValider', True)).OnClick 		:= BValideClick ;

  // On affiche seulement la 1ère ligne des pièces
  SetControlText('E_NUMLIGNE',	'1') ;
  SetControlText('E_NUMLIGNE_',	'1') ;
  SetControlText('E_NUMECHE',		'1') ;



  TToolbarButton97( GetControl('BCHERCHE', True)   ).OnClick := BChercheClick   ;
  // Puisque le 'Plus' des datatype n'est pas repris lorsque '<<Tous>>' est choisi sur E_NATUREPIECE par exemple, il faut reprendre les restrictions
  SetControlText('XX_WHERE',' AND E_CREERPAR<>"DET" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="AC" OR E_NATUREPIECE="FF" OR E_NATUREPIECE="FC") ');

  inherited ;

  InitFormModification;
  InitMsgBox;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
    E_AUXILIAIRE.OnElipsisClick:=AuxiElipsisClick;

  {FQ20226  20.06.07 YMO}  
  TFMul(Ecran).BCherche.Click;
end ;

procedure TOF_CPGENEECRTP.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPGENEECRTP.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPGENEECRTP.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_CPGENEECRTP.E_EXERCICEChange(Sender: TObject);
begin
 	ExoToDates( GetControlText('E_EXERCICE'), TEdit(GetControl('E_DATECOMPTABLE', True)), TEdit(GetControl('E_DATECOMPTABLE_', True)) ) ;
	if ((GetControlText('E_EXERCICE')='')) then
  	BEGIN
    SetControlText('E_DATECOMPTABLE',		stDate1900) ;
    SetControlText('E_DATECOMPTABLE_',	stDate2099) ;
    END ;
end;


procedure TOF_CPGENEECRTP.FListeDblClick(Sender: TObject);
begin
  OuvrirClick ;
end;


procedure TOF_CPGENEECRTP.InitFormModification;
begin

  if ANouveau then
  begin
    E_EXERCICE.Value      := VH^.Encours.Code ;
    E_DATECOMPTABLE.Text  := DateToStr(VH^.Encours.Deb) ;
    E_DATECOMPTABLE_.Text := DateToStr(VH^.Encours.Deb) ;
  end
  else
    if VH^.CPExoRef.Code<>'' then
    begin
      E_EXERCICE.Value      := VH^.CPExoRef.Code ;
      E_DATECOMPTABLE.Text  := DateToStr(VH^.CPExoRef.Deb) ;
      E_DATECOMPTABLE_.Text := DateToStr(VH^.CPExoRef.Fin) ;
    end
    else
    begin   // uniquement si exercice ouvert
      If ExisteSQL('SELECT EX_EXERCICE FROM EXERCICE WHERE EX_EXERCICE="' + VH^.Entree.Code
      +'" AND (EX_ETATCPTA="OUV" OR EX_ETATCPTA="CPR")') then
      begin
          E_EXERCICE.Value      := VH^.Entree.Code ;
          E_DATECOMPTABLE.Text  := DateToStr(V_PGI.DateEntree) ;
          E_DATECOMPTABLE_.Text := DateToStr(V_PGI.DateEntree) ;
      end
      else
      begin
          E_EXERCICE.Value      := VH^.Encours.Code ;
          E_DATECOMPTABLE.Text  := DateToStr(VH^.Encours.Deb) ;
          E_DATECOMPTABLE_.Text := DateToStr(VH^.Encours.Deb) ;
      end;
    end ;

  E_DATEECHEANCE.Text   := stDate1900 ;
  E_DATEECHEANCE_.Text  := stDate2099 ;

  PositionneEtabUser( E_ETABLISSEMENT ) ;
  StopCherche := True; // Pas d'affichage dès l'ouverture

//  E_JOURNAL.Plus:=' AND J_FERME="-" AND J_NATUREJAL<>"ANO" AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ODA" And J_NATUREJAL<>"ANA"';

  //<<Tous>> par défaut
  If (E_JOURNAL.Values.Count > 0) and (E_JOURNAL.Values[0]='') then E_JOURNAL.Value:=E_JOURNAL.Values[0];
//If (E_EXERCICE.Values.Count > 0) and (E_EXERCICE.Values[0]='') then E_EXERCICE.Value:=E_EXERCICE.Values[0];
  If (E_NATUREPIECE.Values.Count > 0) and (E_NATUREPIECE.Values[0]='') then E_NATUREPIECE.Value:=E_NATUREPIECE.Values[0];

  If (E_DEVISE.Values.Count > 0) and (E_DEVISE.Values[0]='') then E_DEVISE.Value:=E_DEVISE.Values[0];
  If (CATEGORIE.Values.Count > 0) and (CATEGORIE.Values[0]='') then CATEGORIE.Value:=CATEGORIE.Values[0];
  If (E_MODEPAIE.Values.Count > 0) and (E_MODEPAIE.Values[0]='') then E_MODEPAIE.Value:=E_MODEPAIE.Values[0];
end;


procedure TOF_CPGENEECRTP.OuvrirClick;
var lDossier : String ;
    lMulQ    : TQuery ;
begin

  if GetDataSet.Bof and GetDataSet.Eof then Exit ;

  // Réaffectation systématique de TheMulQ car vaut nil au 2ème passage
{$IFDEF EAGLCLIENT}
  lMulQ := FEcran.Q.TQ;
  lMulQ.Seek( FListe.Row - 1 );
  lMulQ.PutGridDetail(FListe,true,true,''); // Affichage des en-têtes
{$ELSE}
  lMulQ := FEcran.Q;
{$ENDIF}

  lDossier := V_PGI.SchemaName ;

  TrouveEtLanceSaisie( lMulQ, taConsult, '' ) ;    //GetControlText('E_QUALIFPIECE')

end;

function TOF_CPGENEECRTP.AddFille (T : TOB ; Q : TQuery) : TOB;
var
      TF : Tob;
      {$IFNDEF EAGLCLIENT}
      i : Integer;
      {$ENDIF}
begin
      TF := Tob.Create('ECRITURE_', T, -1);
      {$IFDEF EAGLCLIENT}
      // Renseigne les valeurs
      TF.Dupliquer(Q.CurrentFille,true,true);
      {$ELSE}
      // Renseigne les valeurs
      for i := 0 to Q.FieldCount - 1 do
        TF.AddChampSupValeur(Q.Fields.Fields[i].FieldName, Q.Fields.Fields[i].AsString, False);
      {$ENDIF}

      Result := TF;
end;

procedure TOF_CPGENEECRTP.ValideClick;
var lMulQ : TQuery ;
    i     : integer;
    TobEcr, TobDetail : TOB ;
    M : RMVT ; //YMOO
    ListePieces : TList ;
begin

{$IFDEF EAGLCLIENT}
  lMulQ := FEcran.Q.TQ;
  lMulQ.Seek( FListe.Row - 1 );
{$ELSE}
  lMulQ := FEcran.Q ;
{$ENDIF}

  TobEcr := Tob.Create('ECRITURE_', nil, -1) ;

  If TFmul(Ecran).FListe.AllSelected then
  begin
      LMulQ.First;

      While not LMulQ.EOF do
      begin
         If ExisteTP(LMulQ) then AddFille(TobEcr, LMulQ);
         LMulQ.Next;
      end;
  end
  else
  for i := 0 to FListe.NbSelected - 1 do
  begin
    MoveCur(FALSE);
    FListe.GotoLeBookMark(i);
    {$IFDEF EAGLCLIENT}
    LMulQ.Seek(FListe.Row - 1);
    {$ENDIF}

    If ExisteTP(LMulQ) Then AddFille(TobEcr, LMulQ);
  end;
  ListePieces := TList.Create ;

  for i:=0 to TobEcr.Detail.Count-1 do
  begin
    // Recup TOB
    // O := TOBM( FPiecesGener[i] );
    // Transformation en RMVT

    TobDetail:=TobEcr.Detail[i];

    M := TobToIdent(TobDetail,True);
    // Génération des Pièces sur les TP
    GenerePiecesPayeur(M , '', ListePieces);  {Tierspayeur.pas}

  end;

  If ListePieces.Count>0 then
  begin
    {Pour n'afficher dans le compte rendu que la 1ere ligne, la méthode la + simple
    est de supprimer dans la liste de restitution 1 ligne sur 2, les lignes impaires.
    Pour rappel, cette Tlist contient les lignes de contrepartie et de contrepassation
    générées dans TiersPayeur.pas}
    i:=ListePieces.Count-1;
    While i>0 do
    begin
      ListePieces.Delete(i);
      i:=i-2;
    end;

    VisuPiecesGenere(ListePieces,EcrGen, 18)

  end
  else
    PgiInfo('Aucune écriture traitée');

  VideListe(ListePieces) ;
  ListePieces.Free ;
  TobEcr.Free;

  FiniMove;
end;

Function TOF_CPGENEECRTP.ExisteTP (Q : TQuery) : boolean ;
Var Nat : String ;
BEGIN
Nat := Q.FindField('E_NATUREPIECE').AsString;
Result:=False ;
if Action=taConsult then Exit ;
if Not VH^.OuiTP then Exit ;
//if M.Simul<>'N' then Exit ;
//if M.ANouveau then Exit ;
if Not EstJalFact(Q.FindField('E_JOURNAL').AsString) then Exit ;

if ((NAT<>'FC') and (NAT<>'FF') and (NAT<>'AC') and (NAT<>'AF')) then Exit ;
if ((Nat='FC') or (Nat='AC')) then if VH^.JalVTP='' then
   BEGIN
   PgiInfo('Vous n''avez pas paramétré de journal pour les tiers payeurs, le mécanisme ne sera pas activé') ;
   Exit ;
   END ;
if ((Nat='FF') or (Nat='AF')) then if VH^.JalATP='' then
   BEGIN
   PgiInfo('Vous n''avez pas paramétré de journal pour les tiers payeurs, le mécanisme ne sera pas activé') ;
   Exit ;
   END ;

Result:=True ;

END ;


procedure TOF_CPGENEECRTP.BOuvrirClick(Sender: TObject);
begin
 OuvrirClick ;
end;

procedure TOF_CPGENEECRTP.BValideClick(Sender: TObject);
begin

 If (Fliste.NbSelected=0) and (not Fliste.AllSelected) then Exit ;

 ValideClick ;

 FListe.ClearSelected;

 TFMul(Ecran).BCherche.Click;
end;


procedure TOF_CPGENEECRTP.InitMsgBox;
begin
  MsgBox := THMsgBox.create(FMenuG);
  MsgBox.Mess.Add('0;Validation;Ce traitement va générer les écritures sur Tiers payeurs. Désirez-vous continuer?;Q;YN;Y;;');
  MsgBox.Mess.Add('1;Désirez-vous continuer?;Q;YN;Y;;');
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/10/2006
Modifié le ... :  /  /
Description .. : On n'affiche pas les pièces à l'entrée dans le MUL
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGENEECRTP.BChercheClick(Sender: TObject);
begin
  //YMO 28/11/2006 FQ19227
  {$IFDEF EAGLCLIENT}
  If StopCherche then
  begin // On force le Bcherche à l'entrée du MUL pour l'affichage de tous les en-têtes en CWAS
    SetControlText('E_NUMLIGNE_',	'0') ;
    TFMul(Ecran).BChercheClick( nil );
    SetControlText('E_NUMLIGNE_',	'1') ;
  end;
  {$ELSE}
  HMTrad.ResizeDBGridColumns(FListe) ;
  {$ENDIF}
  If not StopCherche then
    TFMul(Ecran).BChercheClick( nil );
  StopCherche := False;

end;

procedure TOF_CPGENEECRTP.FormKeyDown( Sender : TObject; var Key: Word; Shift: TShiftState );
begin
 inherited;
 if ( csDestroying in Ecran.ComponentState ) then Exit ;

case Key of
  VK_F10 : BValideClick(nil);
  VK_F5 : if FListe.Focused then
           begin
            Key:=0 ;
            FListeDblClick(nil);
           end;
 end;

 if ( Ecran <> nil ) and ( Ecran is  TFMul ) then
    TFMul(Ecran).FormKeyDown(Sender,Key,Shift);

end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGENEECRTP.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

Initialization
  registerclasses ( [ TOF_CPGENEECRTP ] ) ; 
end.
