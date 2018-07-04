{***********UNITE*************************************************
Auteur  ...... : Compta
Créé le ...... : 04/11/2002
Modifié le ... : 09/10/2007
Description .. : Source TOF de la FICHE : CPMULMVT ()
Suite ........ : 
Suite ........ : 09/10/2007 SBO FQ 18431
Mots clefs ... : TOF;CPMULMVT
*****************************************************************}
Unit UTOFCPMULMVT                                                  ;

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
     Ent1,		    // Pour EstMonaieIn
     Ed_Tools,              // Pour le videListe
     TiersPayeur,           // Pour les fonctions xxxTP
     HStatus,               // Pour la barre d'état
     AGLInit,               // TheMulQ
     Zcompte,
     CPANCETREMUL, {JP 15/10/07 : FQ 16149}
     SaisComm;              // Pour les procédures de MAJ des comptes

procedure MultiCritereMvt(Comment : TActionFiche ; Simul : String ; ANouveau : boolean );
procedure DetruitEcritures ( TypeMvt : String ; QueBOR : boolean ) ;
{JP 21/03/07 : Suppression de pièces sur critères}
procedure DetruitSurCritere(Jal, Exo, Nat, Qual, Num, dtC, Etab : string; QueBor : Boolean);

Type
  TOF_CPMULMVT = Class (TOFANCETREMUL)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose ; override ;
  private
    FEcran    : TFMul ;
    Action    : String ;    // 'taAction' ou 'taModif' ou 'taConsult' ou 'taSuppr'
    Simul     :	String ;    // 'N' ou 'S' ou 'R'
    ANouveau  : Boolean ;
    queBOR    : Boolean ;
    FbCharge  : Boolean ;
    sMode     : String ;
    FParams   : RMVT ;

    // Eléments interface
    E_JOURNAL        : THValComboBox ;
    E_NATUREPIECE    : THValComboBox ;
    E_QUALIFPIECE    : THValComboBox ;
    E_EXERCICE       : THValComboBox ;
    E_ETABLISSEMENT  : THValComboBox ;
    E_DATECOMPTABLE  : THEdit ;
    E_DATECOMPTABLE_ : THEdit ;
    E_DATECREATION   : THEdit ;
    E_DATECREATION_  : THEdit ;
    E_DATEECHEANCE   : THEdit ;
    E_DATEECHEANCE_  : THEdit ;
    XX_WHERE         : THEdit ;
    XX_WHEREBOR      : THEdit ;
    XX_WHEREGC       : THEdit ;
    Rune             : TCheckBox ;
  {$IFDEF EAGLCLIENT}
    FListe : THGrid ;
  {$ELSE}
    FListe : THDBGrid ;
  {$ENDIF}
    // Gestion Suppression
    ListePieces : TList ;
    ListeTP     : HTStrings ;
    NowFutur    : TDateTime ;
    ExistPasDet : boolean ;
    OkVuPiecesDetruites  : boolean ;
    FBoSaisieParam : boolean ; // Test saisie paramétrable

    {$IFDEF COMPTA}
    {JP 15/05/06 : Si on gère la suppression en MultiSoc, il faudra faire le test sur
                   chaque ligne et ne pas gérer un Booléen général}
    FBoGestionBap : Boolean;
    AttenteLettrable : Boolean; // ajout me
    {$ENDIF COMPTA}
{b fb 02/05/2006}
    Engagement : boolean;
    Prevision : boolean;
{e fb 02/05/2006}

    // Evènements fiche
    procedure AfterShow;
    procedure RUneClick(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure E_DEVISEChange(Sender: TObject);
    procedure BChangeTauxClick(Sender: TObject);
    procedure BTobVClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure ValideClick ;
    procedure ValideSupprClick ;
    procedure BDetruitesClick( Sender : TObject ) ;
    procedure BZoomPieceClick( Sender : TObject ) ;
    procedure BValideClick ( Sender : TObject ) ;
// Autres procédures
    procedure RetoucheCritere ;
    procedure InitCriteres ;
    procedure AttribTitre ;
    procedure AttribTitreSuppr ;
    procedure AttribListe;
    function  DetruitSelection : boolean ;
    procedure DetruitLesEngagements ( TPiece : TList ) ;  //fb 02/05/2006
    procedure DetruitLesEcritures ( TPiece : TList ) ;  // ok
    procedure DetruitLesEcrTP ;                         // ok
    procedure DetruitLaPiece ;                          // ok
    procedure DetruitLesSections ( Q1 : TQuery ; GeneTypeExo : TTypeexo ) ;  // ok
    procedure InverseDetruitSoldes ( Q1 : TQuery ; GeneTypeExo : TTypeExo ) ;
    Function  InverseSoldesSections ( Q1 : TQuery ; GeneTypeExo : TTypeExo ) : longint ;
    function  JePeuxDetruire ( Q1 : TQuery ) : boolean ;
    function  UpdateSet ( NumL : integer ; EcartChange : boolean ) : String ;
    Procedure ChargeComboExo ;
    Procedure InitFormSuppression ;
    Procedure InitFormModification ;
    Procedure VirerChoixTous( vCombo : THValComboBox ) ;
    procedure DetruitLesAnoDyna( vTOB : TOB ) ;
    {$IFDEF COMPTA}
    function CanDeleteBap(Nat, Qual, NumPiece, UsDateCpt, Exo, Jal : string) : Boolean;
    {$ENDIF COMPTA}
    // FQ18445 : Tout sélectionner sur suppression d'écritures de simul et de situ
    function  DetruitSelectionSimul : boolean ;
    procedure DetruitAllSimul ;
    function  JePeuxDetruireAllSimul : Boolean ;
    procedure AuxiElipsisClick(Sender : TObject);
  end ;

Implementation

uses
{$IFDEF MODENT1}
  ULibExercice,
  CPTypeCons,
{$ENDIF MODENT1}

{$IFDEF EAGLCLIENT}
{$ELSE}
{$ENDIF}  // Bouton chgt de taux actif en eAGL YMO 01/06
  RepDevEur, // ChangeLeTauxDevise
  UlibEcriture, Constantes,
  {$IFDEF COMPTA}
  uLibBonAPayer,
  {$ENDIF COMPTA}
  DelVisuE  // Pour les procédures d'affichage des pieces détruites
  , UTofMulParamGen; {13/04/07 YMO F5 sur Auxiliaire }

// Message d'erreurs
Const MessageListe : Array[0..13] of String =
					( 'Visualisation des écritures',
            'Modification des écritures',
            'courantes',
            'de simulation',
            'de prévision',
            'de situation',
            'de révision',
            'd''à-nouveaux',
            'Origine',
            'Vous n''avez pas renseigné de journal ni de comptes. La sélection peut être importante. Confirmez-vous l''analyse ?',
            'Suppresion des écritures courantes',
            'Suppresion des écritures en mode bordereau',
            'IFRS',
            'd''engagement'
           );  //fb 10/03/2005

Const MessageListeSuppr : Array[0..16] of String =
          ( 'Suppression des écritures courantes',
            'Suppression des écritures de simulation',
            'Suppression des écritures de situation',
            'Suppression des écritures de prévision',
            '4;?caption?;Vous n''avez sélectionné aucune écriture;E;O;O;O;',
            '5;?caption?;Confirmez-vous la suppression des écritures sélectionnées ?;Q;YN;N;N;',
            'DETRUITE LE',
            'PAR',
            '8;?caption?;Voulez-vous voir la liste des écritures détruites ?;Q;YN;N;N;',
            'ATTENTION! Certaines écritures en cours de traitement par un autre utilisateur n''ont pas été détruites',
            'Suppression des écritures de révision',
            '11;?caption?;Certaines écritures;E;O;O;O;',
            {$IFDEF COMPTA}
            '12;?caption?;Des écritures validées, pointées, lettrées, avec des immobilisation ou'#13 +
                  'un bon à payer validé n''ont pas été détruites.;E;O;O;O;',
            {$ELSE}
            '12;?caption?;Des écritures validées, pointées, lettrées ou avec immobilisations n''ont pas été détruites;E;O;O;O;',
            {$ENDIF COMPTA}
            'Suppression des écritures en mode bordereau',
            'Suppression des écritures IFRS',
            'Suppression des écritures d''engagement', //fb 02/05/2006
            'La sélection contient des écritures validées ou portant sur un exercice clos. La suppression est impossible.'
           );

//==============================================================================

procedure MultiCritereMvt(Comment : TActionFiche ; Simul : String ; ANouveau : boolean );
var Args : String;
{$IFDEF CCSTD}
{$ELSE}
  lSt       : string ;
  lStQualif : string ;
  {$IFDEF COMPTA}
    M : RMVT ;
  {$ENDIF}
  lBoSaisieParam : Boolean ;
{$ENDIF}
begin
{$IFDEF CCSTD}

{$ELSE}
{$IFDEF ENTREPRISE}
If _GetParamSocSecur('SO_CPANODYNA',false,TRUE) Then ;
{$ENDIF}

  // Gestion appel depuis journal centralisateur
  //  --> conflit avec l'appel pour la saisie paramétrable
  // SBO 15/11/2006
  lSt       := Simul ;
  lStQualif := READTOKENST(lSt) ;
  lBoSaisieParam := ( pos( 'SP', lStQualif ) > 0 ) ;
  if lBoSaisieParam then
    lStQualif := copy( lStQualif, 3, 1 ) ;

  {$IFDEF COMPTA}
  if Comment=taCreat then
  BEGIN
    FillChar(M,Sizeof(M),#0) ;
    M.Simul         := lStQualif ;
    M.Jal           := READTOKENST(lSt) ;
    lSt             := READTOKENST(lSt) ;
    if lSt <> ''
      then M.DateC 	:= StrToDate(lSt)
      else M.DateC 	:= V_PGI.DateEntree ;
    M.ParCentral    := length(M.Jal) > 0 ;
    M.CodeD 		    := V_PGI.DevisePivot ;
    M.TauxD 		    := 1 ;
    M.DateTaux	    := M.DateC ;
    M.Valide		    := False ;
    M.Etabl		      := VH^.ETABLISDEFAUT ;
    M.ANouveau	    := ANouveau ;
    if lBoSaisieParam
      then LanceSaisieParam(nil, taCreat, M)
      else LanceSaisie(Nil,taCreat,M) ;
    Exit ;
  END;
  {$ENDIF}
{  else
   if ANouveau
   	then TypeAction	:= 'H'
    else TypeAction	:= Simul ;
}
  if Comment = taModif
    then Args := 'taModif'
    else Args := 'taConsult';
  Args := Args + ';' + lStQualif ;
  if ANouveau
    then Args := Args + ';TRUE'
    else Args := Args + ';FALSE';
  if lBoSaisieParam then
    Args := Args + ';SAISIEPARAM';
{$ENDIF}

  AGLLanceFiche('CP','CPMULMVT','','',Args);
end;

procedure DetruitEcritures ( TypeMvt : String ; QueBOR : boolean ) ;
var lStArgs 		: String;
begin

  if _Blocage(['nrCloture'],True,'nrAucun') then Exit ; // FQ 17425

  lStArgs := 'taSuppr;' + TypeMvt ;
  if QueBOR
    then lStArgs := lStArgs + ';TRUE'
    else lStArgs := lStArgs + ';FALSE';

  AGLLanceFiche('CP', 'CPMULMVT', '', '', lStArgs);
end;

{JP 21/03/07 : Suppression de pièces sur critères
{---------------------------------------------------------------------------------------}
procedure DetruitSurCritere(Jal, Exo, Nat, Qual, Num, dtC, Etab : string; QueBor : Boolean);
{---------------------------------------------------------------------------------------}
var
  lStArgs : string;
  Range   : string;
begin
  if _Blocage(['nrCloture'],True,'nrAucun') then Exit;

  lStArgs := 'taSuppr;' + Qual;
  if QueBOR then lStArgs := lStArgs + ';TRUE'
            else lStArgs := lStArgs + ';FALSE';

  Range := 'E_JOURNAL=' + Jal + ';E_EXERCICE=' + Exo + ';E_NATUREPIECE=' + Nat + ';E_QUALIFPIECE=' +
           {Pour la date comptable, je précise la fourchette, mais ce n'est pas nécessaire pour le numéro de pièce}
           Qual + ';E_NUMEROPIECE=' + Num + ';E_DATECOMPTABLE=' + dtC  + ';E_DATECOMPTABLE_=' + dtC + ';E_ETABLISSEMENT=' + Etab + ';';
  AGLLanceFiche('CP', 'CPMULMVT', Range, '', lStArgs);
end;

//==============================================================================
procedure TOF_CPMULMVT.OnUpdate ;
begin
  RetoucheCritere ;
  if (Action='taModif') then
    begin
    if (GetControlText('E_DEVISE')<>'')
      and (GetControlText('E_Devise')<>V_PGI.DevisePivot)
      and (GetControlText('E_Devise')<>V_PGI.DeviseFongible)
      {and (not EstMonnaieIn(GetControlText('E_Devise')))}
      and (not ANouveau)
      then SetControlVisible('BChangeTaux',TRUE)
      else SetControlVisible('BChangeTaux',FALSE) ;
    end ;
  Inherited ;
end ;

procedure TOF_CPMULMVT.OnArgument (S : String ) ;
Var stANouveau : String ;
begin
  TFMul(Ecran).OnAfterFormShow := AfterShow;

  // Récupération des arguments
  if Trim(S) <> '' then
    Action 	:= ReadTokenSt(S) ;
  if Trim(S) <> '' then
    Simul 	:= ReadTokenSt(S) ;
  if Trim(S) <> '' then
    stANouveau:= ReadTokenSt(S) ;
  if stANouveau = 'TRUE' then
    ANouveau := True
  else
    ANouveau := False ;

  FbCharge := FALSE ; // ??
  sMode	   := '-' ;   // ??
  SetControlProperty('E_VALIDE','State',cbGrayed);

  if Trim(S) <> ''
    then FBoSaisieParam := ( ReadTokenSt(S) = 'SAISIEPARAM' )
    else FBoSaisieParam := False ;

  // Attention, pour le mode taSuppr, ANouveau est utilise pour stocke QueBor
  if (Action='taSuppr') then
    begin
    QueBor := ANouveau ;
    ANouveau := False ;
    SetControlVisible('E_VALIDE',False) ;
    TToolBarButton97(GetControl('Bouvrir',true)).Glyph:=TToolBarButton97(getcontrol('BJUSTEPOURICONE',true)).Glyph; //SG6 23/11/04 FQ 14202
    TToolBarButton97(GetControl('Bouvrir',true)).NumGlyphs:=2; //SG6 23/11/04 FQ 14202
    SetControlProperty('BOuvrir','Hint','Supprimer');

    {$IFDEF COMPTA}
    FBoGestionBap := ExisteTypeVisa('');
    {$ENDIF COMPTA}
    end ;

  // Récup interface
  FEcran := TFMul(Ecran) ;
{$IFDEF EAGLCLIENT}
  FListe  := THGrid( GetControl('FListe',True) ) ;
  TheMulQ := FEcran.Q.TQ;
{$ELSE}
  FListe  := THDBGrid( GetControl('FListe',True)) ;
  TheMulQ := FEcran.Q;
{$ENDIF}


  E_JOURNAL        := THValComboBox(GetControl('E_JOURNAL', True)) ;
  E_NATUREPIECE    := THValComboBox(GetControl('E_NATUREPIECE', True)) ;
  E_QUALIFPIECE    := THValComboBox(GetControl('E_QUALIFPIECE', True)) ;
  E_EXERCICE       := THValComboBox(GetControl('E_EXERCICE', True)) ;
  E_ETABLISSEMENT  := THValComboBox(GetControl('E_ETABLISSEMENT', True)) ;
  E_DATECOMPTABLE  := THEdit(GetControl('E_DATECOMPTABLE', True))  ;
  E_DATECOMPTABLE_ := THEdit(GetControl('E_DATECOMPTABLE_', True)) ;
  E_DATECREATION   := THEdit(GetControl('E_DATECREATION', True))  ;
  E_DATECREATION_  := THEdit(GetControl('E_DATECREATION_', True)) ;
  E_DATEECHEANCE   := THEdit(GetControl('E_DATEECHEANCE', True)) ;
  E_DATEECHEANCE_  := THEdit(GetControl('E_DATEECHEANCE_', True)) ;
  XX_WHERE         := THEdit(GetControl('XX_WHERE', True)) ;
  {JP 29/07/05 : c'est en traitant la FQ 15124, que je suis tombé sur le problème et je pense que la
                 FQ 15259 était lié au fait que les trois commposants XX_WHERE pointaient sur le
                 même composant}
  XX_WHEREBOR      := THEdit(GetControl('XX_WHEREBOR', True)) ;
  XX_WHEREGC       := THEdit(GetControl('XX_WHEREGC', True)) ;
  Rune             := TCheckBox(GetControl('Rune', True)) ;

	// Réaffectation des évènements
  E_EXERCICE.OnChange 	                                  := E_EXERCICEChange ;
  THValComboBox(GetControl('E_DEVISE', True)).OnChange 	  := E_DEVISEChange ;
  FListe.OnDblClick 					  := FListeDblClick ;
  RUne.onClick 						  := RUneClick ;
  TButton(GetControl('BChangeTaux', True)).OnClick 	  := BChangeTauxClick ;
  TButton(GetControl('BTobV', True)).OnClick 		  := BTobVClick ;
  TButton(GetControl('BDetruites', True)).OnClick 	  := BDetruitesClick ;
  TButton(GetControl('BZoomPiece', True)).OnClick 	  := BZoomPieceClick ;
  TButton(GetControl('BOuvrir', True)).OnClick 		  := BValideClick ;
  // fiche 19708
  if GetCPExoRef.Code <> '' then
  begin
    SetControlText('E_EXERCICE', GetCPExoRef.Code);
    SetControlText('E_DATECOMPTABLE', DateToStr(GetCPExoRef.Deb));
    SetControlText('E_DATECOMPTABLE_', DateToStr(GetCPExoRef.Fin));
  end
  else
  begin
    SetControlText('E_EXERCICE', GetEntree.Code);
    SetControlText('E_DATECOMPTABLE', DateToStr(V_PGI.DateEntree));
    SetControlText('E_DATECOMPTABLE_', DateToStr(V_PGI.DateEntree));
  end;
  SetControlText('E_DATEECHEANCE', StDate1900);
  SetControlText('E_DATEECHEANCE_', StDate2099);

  // Visibilité des boutons
  // FQ18745 : Suppression en masse pour écritures de simul
  SetControlVisible('BSelectAll', (Action='taSuppr') and ((Simul='S') or (Simul='U')) ) ;
//  SetControlVisible('BSelectAll', False) ;
  if Action='taSuppr' then
    SetControlVisible('BChangeTaux', False) ;
  SetControlVisible('BDetruites', (Action='taSuppr')) ;
  SetControlVisible('BZoomPiece', (Action='taSuppr')) ;
  SetControlVisible('BTOBV',True) ;   // FQ 17317


  // Restriction communes aux différents mode d'utilisation
  E_QUALIFPIECE.Value   := Simul ;
  E_QUALIFPIECE.Enabled := False ;

  // Init suivant utilisationd e la fiche
{b fb 02/05/2006}
  Engagement := (Simul='p');
  Prevision  := (Simul='P');
  if Engagement then
    Simul:='P';
{e fb 02/05/2006}
  if (Action='taSuppr')
    then InitFormSuppression
    else InitFormModification ;

  inherited ;
  {JP 06/10/05 : Gestion des filtres et des listes avec la nouvelle gestion de l'agl}
  AttribListe;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
    THEdit(GetControl('E_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;
end ;

procedure TOF_CPMULMVT.InitCriteres;
begin
  if VH^.Precedent.Code<>''
    then E_DATECOMPTABLE.Text := DateToStr(VH^.Precedent.Deb)
    else E_DATECOMPTABLE.Text := DateToStr(VH^.Encours.Deb) ;
  E_DATECOMPTABLE_.Text := DateToStr(V_PGI.DateEntree) ;
end;

procedure TOF_CPMULMVT.AttribTitre;
Var Ind     : integer ;
    stValue : string;
BEGIN
  Ind:=0 ;

  // HelpContext
  stValue := E_QUALIFPIECE.Value ;
  if ANouveau then
    BEGIN
    Ind:=7 ;
    if Action='taConsult' then
      Ecran.HelpContext:=7727000 ;
    END
  else if stValue='N' then
    BEGIN
    Ind:=2 ;
    if Action='taConsult' then
      begin
      if FBoSaisieParam then Ecran.HelpContext:=7246200 // FQ 16386 SBO 11/10/2005
                        else Ecran.HelpContext:=7256000
      end
    else
      begin
      if FBoSaisieParam then Ecran.HelpContext:=7246300 // FQ 16386 SBO 11/10/2005
                        else Ecran.HelpContext:=7259000 ;
      end ;
    END
  else if stValue='S' then
    BEGIN
    Ind:=3 ;
    if Action='taConsult'
      then Ecran.HelpContext:=7274000
      else Ecran.HelpContext:=7280000 ;
    END
{b fb 10/03/2005}
  else if Engagement then
    begin
    Ind:=13
    end
{e fb 10/03/2005}
  else if stValue='P' then
    BEGIN
    Ind:=4 ;
    if Action='taConsult'
      then Ecran.HelpContext:=7310000
      else Ecran.HelpContext:=7316000 ;
    END
  else if stValue='U' then
    BEGIN
    Ind:=5 ;
    if Action='taConsult'
      then Ecran.HelpContext:=7295000
      else Ecran.HelpContext:=7301000 ;
    END
  else if stValue='R' then
    BEGIN
    Ind:=6 ;
    if Action='taConsult'
      then Ecran.HelpContext:=7658000
      else Ecran.HelpContext:=7661000 ;
    END
  else if stValue='I' then // MODIF IFRS
      BEGIN
      Ind:=12 ;
      if Action='taConsult'
        then Ecran.HelpContext:=7249200
        else Ecran.HelpContext:=7249300 ;
      END ;

  // Saisie paramétrable
  if FBoSaisieParam then
    if Action='taConsult'
      then Ecran.HelpContext:=7246200   // FQ 16386 SBO 11/10/2005
      else Ecran.HelpContext:=7246300 ; // FQ 16386 SBO 11/10/2005

  // Titre
  if Ind>0 then Ecran.Caption := Ecran.Caption + ' ' + MessageListe[Ind] ;
  if FBoSaisieParam then Ecran.Caption := Ecran.Caption + ' (via saisie paramétrable)';
  UpdateCaption(Ecran) ;
end;

procedure TOF_CPMULMVT.RetoucheCritere;
begin
  if (GetControlText('E_GENERAL')<>'') Or (GetControlText('E_AUXILIAIRE')<>'')
    then SetControlChecked('RUne',FALSE) ;
end;

procedure TOF_CPMULMVT.BChangeTauxClick(Sender: TObject);
var TOBOrig : Tob ;
    TobL : Tob;
    i     : Integer ;
    lQEcr : TQuery ;
ReqEcriture : String;
begin
{$IFDEF EAGLCLIENT}
{$ELSE}
{$ENDIF} // Bouton chgt de taux actif en eAGL YMO 01/06

 if FListe.nbSelected > 0 then
 begin
     TOBOrig := TOB.Create( 'TOB_ORIGINE', nil, -1 ) ;

     For i:=0 to FListe.NbSelected-1 do
     begin
        FListe.GotoLeBookmark(i) ;

        ReqEcriture := 'SELECT ECRITURE.* FROM ECRITURE '
                 + 'WHERE E_JOURNAL="' + GetField('E_JOURNAL')  + '" AND '
                 + 'E_EXERCICE="' + GetField('E_EXERCICE') + '" AND '
                 + 'E_DATECOMPTABLE="' + UsDateTime(GetField('E_DATECOMPTABLE')) + '" AND '
                 + 'E_NUMEROPIECE='     + IntToStr(GetField('E_NUMEROPIECE')) + ' AND '
                 + 'E_NUMLIGNE=' + IntToStr(GetField('E_NUMLIGNE')) + ' AND '
                 + 'E_NUMECHE=' + IntToStr(GetField('E_NUMECHE')) ;

        TOBL := TOB.Create('ECRITURE', TOBOrig, -1 ) ;
        lQEcr := OpenSQL( ReqEcriture, True ) ;
        TOBL.SelectDB( '', lQEcr ) ;
        Ferme( lQEcr ) ;
     end ;

     ChangeLeTauxDevise(TOBOrig, TaModif) ;
     FreeAndNil(TOBOrig) ;
 end
 else
     PGIInfo(TraduireMemoire('Veuillez sélectionner une pièce'),'Gestion du taux');
//  FreeAndNil(TOBL) ;

end;

procedure TOF_CPMULMVT.BTobVClick(Sender: TObject);
Var WhereSQL : String ;
begin
  if ((GetControlText('E_JOURNAL')='') and (GetControlText('E_GENERAL')='')
                                       and (GetControlText('E_AUXILIAIRE')='') ) then
    if PGIAsk(MessageListe[9], Ecran.Caption) <> mrYes
      then Exit ;

  WhereSQL := RecupWhereCritere(TPageControl(GetControl('Pages', True))) ;
  AGLLanceFiche('CP','CPCONSULTREVIS','','',WhereSQL) ;

end;

procedure TOF_CPMULMVT.E_DEVISEChange(Sender: TObject);
begin
  If (Action='taModif') then
    begin
    If (GetControlText('E_DEVISE')<>'')
        and (GetControlText('E_Devise')<>V_PGI.DevisePivot)
        and (GetControlText('E_Devise')<>V_PGI.DeviseFongible)
        {and (Not EstMonnaieIn(GetControlText('E_Devise')))}
        and (Not ANouveau) Then
     begin
       SetControlVisible('BChangeTaux', TRUE );
       {$IFDEF EAGLCLIENT}
          FListe.Multiselect:=TRUE;
       {$ELSE}
          FListe.Multiselection:=TRUE; //YMO 16/12/05 changement taux en série
       {$ENDIF}
     end
    Else
     begin
       SetControlVisible('BChangeTaux', FALSE ) ;
       {$IFDEF EAGLCLIENT}
            FListe.Multiselect:=FALSE;
       {$ELSE}
            FListe.Multiselection:=FALSE;
       {$ENDIF}
     end;
    end ;
end;

procedure TOF_CPMULMVT.E_EXERCICEChange(Sender: TObject);
begin
	ExoToDates( GetControlText('E_EXERCICE'), TEdit(GetControl('E_DATECOMPTABLE', True)), TEdit(GetControl('E_DATECOMPTABLE_', True)) ) ;
  // FQ14815 SG6 03/11/2004 Initialisation des dates si on choisit <<Tous>> dans la combo exercice
	if ((GetControlText('E_EXERCICE')='')) then
  	BEGIN
    SetControlText('E_DATECOMPTABLE',		stDate1900) ;
    SetControlText('E_DATECOMPTABLE_',	stDate2099) ;
    END ;
end;

procedure TOF_CPMULMVT.RUneClick(Sender: TObject);
begin
  if GetCheckBoxState('RUne') = cbChecked then
    begin
    SetControlText('E_NUMECHE',		'1') ;
    SetControlText('E_NUMLIGNE',	'1') ;
    SetControlText('E_NUMLIGNE_',	'1') ;
    end
  else
    begin
    SetControlText('E_NUMECHE',		'9999') ;
    SetControlText('E_NUMLIGNE',	'0') ;
    SetControlText('E_NUMLIGNE_',	'9999') ;
    end ;
end;

procedure TOF_CPMULMVT.FListeDblClick(Sender: TObject);
begin
  ValideClick ;
end;


procedure TOF_CPMULMVT.ChargeComboExo ;
var QLoc : TQuery ;
begin
  // FQ 18881 : SBO 25/09/2006 Ajout des exo clo provisoirement
  QLoc := OpenSql('Select EX_EXERCICE,EX_LIBELLE From EXERCICE Where (EX_ETATCPTA="OUV" OR EX_ETATCPTA="CPR") Order By EX_DATEDEBUT', True) ;
  E_EXERCICE.Values.Clear ;
  E_EXERCICE.Items.Clear ;
  while not QLoc.Eof do
    begin
    E_EXERCICE.Values.Add( QLoc.FindField('EX_EXERCICE').AsString ) ;
    E_EXERCICE.Items.Add(  QLoc.FindField('EX_LIBELLE').AsString ) ;
    QLoc.Next ;
    end ;
  Ferme(QLoc) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 03/01/2007
Modifié le ... :   /  /    
Description .. : - LG - 03/01/2007 - FB 19398 - les balances dyna n'etaient 
Suite ........ : pas mise a jour si les ano dyna n'etaient pas cochées
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULMVT.DetruitLaPiece;
Var Q1           : TQuery ;
    StE, StSQL   : String ;
    GeneTypeExo  : TTypeExo ;
    M            : RMVT ;
    TPiece       : TList ;
    P            : P_MV ;
    O            : TOBM ;
    DDEb,DFin,DD : TDateTime ;
    lBoDelOk     : Boolean ;
    lTOB         : TOB ;   
begin
  TPiece  := Tlist.Create ;
  lTOB    := TOB.Create('',nil,-1) ;
  O       := Nil ;
  DD      := GetField('E_DATECOMPTABLE') ;
  DDeb    := DebutdeMois(DD) ;
  DFin    := FindeMois(DD) ;

  StSQL   := 'SELECT ' + GetSelectAll('E', True ) + ', E_BLOCNOTE FROM ECRITURE WHERE E_JOURNAL="' + GetField('E_JOURNAL') + '"'
              + ' AND E_EXERCICE="' + QuelExo( DateToStr(DD) ) + '"'
              + ' AND E_DATECOMPTABLE>="' + USDATETIME(DDeb) + '"'
              + ' AND E_DATECOMPTABLE<="' + USDATETIME(DFin) + '"'
              + ' AND E_QUALIFPIECE="' + E_QUALIFPIECE.Value + '"'
              + ' AND E_NUMEROPIECE=' + IntToStr( GetField('E_NUMEROPIECE') ) ;

  Q1      := OpenSQL( StSQL , True) ;
  if not Q1.EOF then
    begin
    if JePeuxDetruire(Q1) then
      begin
      Q1.First ;
      GeneTypeExo := teEncours ;
      O :=TOBM.Create(EcrGen,'',True) ;
      O.ChargeMvt(Q1) ;
      // PRE-TRAITEMENT SUR SUPPRESSION ECRITURE...
      lBoDelOk := OnDeleteEcritureTob( O, taModif, [cEcrCompl] ) ;
      // SI TRAITEMENT OK, on continue
      if lBoDelOk then
        begin
        {$IFNDEF EAGLCLIENT}
        lTOB.LoadDetailDB('ECRITURE','','',Q1,true);
        {$ENDIF EAGLCLIENT}
        StE := Q1.FindField('E_EXERCICE').AsString ;
        if StE=VH^.Encours.Code
          then GeneTypeExo:=teEncours
          else if StE=VH^.Suivant.Code
                 then GeneTypeExo:=teSuivant ;
        M := MvtToIdent(Q1,fbGene,False) ;
        InverseDetruitSoldes(Q1,GeneTypeExo) ;
        P := P_MV.Create ;
        P.R := M ;
        TPiece.Add(P) ;
        end
      else
        begin
        FreeAndNil( O ) ;
        ExistPasDet := True ;
        end ;
      end
    else
      begin
      ExistPasDet := True ;
      end ;
    end ;
  Ferme(Q1) ;
  if TPiece.Count>0 then
    begin
    DetruitLesEngagements(TPiece) ; //fb 02/05/2006
    DetruitLesEcritures(TPiece) ;
    DetruitLesEcrTP ;
    end ;
  if ((V_PGI.IoError=oeOK) and (O<>Nil)) then
   begin
    ListePieces.Add(O) ;
    DetruitLesAnoDyna(lTOB) ;
   end
    else
     if O<>Nil then
      O.Free ;

  VideListe(TPiece) ;
  TPiece.Free ;
  lTOB.Free ;

end;

{b fb 02/05/2006}
procedure TOF_CPMULMVT.DetruitLesEngagements(TPiece: TList);
Var
  M           : RMVT ;
  i           : integer ;
  S           : String ;
  DDeb,DFin   : TDateTime ;
begin
  if not Engagement then
    Exit;
  if TPiece.Count <= 0 then
    Exit;
  for i:=0 to TPiece.Count-1 do begin
    M := P_MV(TPiece[i]).R ;
    if M.Simul='P' then begin
      DDeb := DebutdeMois(M.DateC) ;
      DFin := FindeMois(M.DateC) ;
      s    := 'CEN_JOURNAL="' + M.JAL + '"'
              + ' AND CEN_EXERCICE="' + M.EXO + '"'
              + ' AND CEN_DATECOMPTABLE>="' + UsDateTime(DDeb) + '"'
              + ' AND CEN_DATECOMPTABLE<="' + UsDateTime(DFin)+'"'
              + ' AND CEN_NUMEROPIECE=' + InttoStr(M.Num)
              + ' AND CEN_QUALIFPIECE="'+M.Simul+'"' ;
       ExecuteSQL('DELETE FROM CENGAGEMENT WHERE ' + S ) ;
      end ;
    end ;
end;
{e fb 02/05/2006}
procedure TOF_CPMULMVT.DetruitLesEcritures(TPiece: TList);
Var M           : RMVT ;
    i           : integer ;
    S           : String ;
    DDeb,DFin   : TDateTime ;
    EcartChange : boolean ;
begin
  if TPiece.Count <= 0 then Exit ;
  for i:=0 to TPiece.Count-1 do
    begin
    M := P_MV(TPiece[i]).R ;
    if not QueBOR
      then S := WhereEcriture(tsGene,M,False)
      else
        begin
        DDeb := DebutdeMois(M.DateC) ;
        DFin := FindeMois(M.DateC) ;
        s    := 'E_JOURNAL="' + M.JAL + '"'
                + ' AND E_EXERCICE="' + M.EXO + '"'
                + ' AND E_DATECOMPTABLE>="' + UsDateTime(DDeb) + '"'
                + ' AND E_DATECOMPTABLE<="' + UsDateTime(DFin)+'"'
                + ' AND E_NUMEROPIECE=' + InttoStr(M.Num)
                + ' AND E_QUALIFPIECE="'+M.Simul+'"' ;
        end ;
    if ( E_QUALIFPIECE.Value = 'N' ) or ( E_QUALIFPIECE.Value = 'I' ) then // Modif IFRS
      begin
      if ((QueBOR) and (ctxPCL in V_PGI.PGIContexte)) then
        ExecuteSQL('DELETE FROM ECRITURE WHERE ' + S )
      else
        begin
        EcartChange := (M.Nature='ECC') ;
        ExecuteSQL('DELETE FROM ECRITURE WHERE ' + S + ' AND (E_NUMLIGNE>2 OR E_NUMECHE>1)') ;
        if ExecuteSQL('UPDATE ECRITURE SET ' + UpdateSet(1,EcartChange)
                                   + ' WHERE ' + S + ' AND E_NUMLIGNE=1')  <> 1 then
          begin
          V_PGI.IOError := oeSaisie ;
          Exit ;
          end ;
        if ExecuteSQL('UPDATE ECRITURE SET ' + UpdateSet(2,EcartChange)
                                   + ' WHERE ' + S + ' AND E_NUMLIGNE=2')  <> 1 then
          begin
          V_PGI.IOError := oeSaisie ;
          Exit ;
          end ;
        end ;
      end
    else
       ExecuteSQL('DELETE FROM ECRITURE WHERE ' + S ) ;
    end ;
end;

procedure TOF_CPMULMVT.DetruitLesEcrTP;
var i           : integer ;
    SQL,PieceTP : String ;
begin
  if Not VH^.OuiTP then Exit ;
  {$IFNDEF CCS3}
  if ListeTP.Count <= 0 then Exit ;
  if QueBOR then Exit ;
  for i:=0 to ListeTP.Count-1 do
    begin
    PieceTP := ListeTP[i] ;
    if PieceTP = '' then Continue ;
    SQL := 'SELECT E_ETATLETTRAGE FROM ECRITURE WHERE ' + WherePieceTP(PieceTP,True)
            + ' AND E_ETATLETTRAGE<>"RI" AND E_ETATLETTRAGE<>"AL"' ;
    if not ExisteSQL( SQL )
      then SupprimePieceTP(PieceTP) ;
    end ;
  {$ENDIF}
end;

procedure TOF_CPMULMVT.DetruitLesSections(Q1: TQuery; GeneTypeExo: TTypeexo);
Var M  : RMVT ;
    QS : TQuery ;
    S  : String ;
    Nb : integer ;
    DDeb,DFin : TDateTime ;
begin
  if Q1.FindField('E_ANA').AsString <> 'X' then Exit ;
  M := MvtToIdent( Q1, fbGene, False) ;
  if Not QueBOR
    then S := WhereEcriture(tsAnal,M,False)
    else
      begin
      DDeb := DebutdeMois(M.DateC) ;
      DFin := FindeMois(M.DateC) ;
      s := 'Y_JOURNAL="' + M.JAL + '"'
           + ' AND Y_EXERCICE="' + M.EXO + '"'
           + ' AND Y_DATECOMPTABLE>="' + UsDateTime(DDeb) + '"'
           + ' AND Y_DATECOMPTABLE<="' + UsDateTime(DFin) + '"'
           + ' AND Y_NUMEROPIECE=' + InttoStr(M.Num)
           + ' AND Y_QUALIFPIECE="' + M.Simul + '"' ;
      end ;
  QS := OpenSQL('SELECT ' + GetSelectAll('Y', True) + ', Y_BLOCNOTE FROM ANALYTIQ WHERE ' + S , True ) ;
  nb := 0 ;
  if not QS.EOF
    then Nb := InverseSoldesSections(QS,GeneTypeExo) ;
  Ferme(QS) ;
  if E_QUALIFPIECE.Value = 'N' then
    begin
    if ExecuteSQL( 'DELETE FROM ANALYTIQ WHERE ' + S ) <> Nb
      then V_PGI.IOError:=oeSaisie ;
    end
  else
    ExecuteSQL( 'DELETE FROM ANALYTIQ WHERE ' + S ) ;
end;

function TOF_CPMULMVT.DetruitSelection: boolean;
Var i,NbD : integer ;
begin
  Result  := False ;
  NbD     := FListe.NbSelected ;
  if (NbD<=0) then
    begin
    HShowMessage(MessageListeSuppr[4],'','') ;
    Exit ;
    end;
  if HShowMessage(MessageListeSuppr[5],'','') <> mrYes then
    begin
    FListe.ClearSelected ;
    Exit ;
    end ;
//  Application.ProcessMessages ;
  ExistPasDet := False ;
  NowFutur    := NowH ;

  InitMove(NbD,'') ;
  for i:=0 to NbD-1 do
    begin
    FListe.GotoLeBookMark(i) ;
{$IFDEF EAGLCLIENT}
    FEcran.Q.TQ.Seek (FListe.Row-1) ;
{$ENDIF}
    MoveCur(FALSE) ;
    if Transactions(DetruitLaPiece,3)<>oeOK then
      begin
      MessageAlerte( MessageListeSuppr[9] ) ;
      FiniMove ;
      Exit ;
      end ;
    end ;
  FListe.ClearSelected ;
  FiniMove ;

  Result := True ;
  if ExistPasDet
    then HShowMessage(MessageListeSuppr[12],'','') ;
end;

procedure TOF_CPMULMVT.InverseDetruitSoldes(Q1: TQuery; GeneTypeExo: TTypeExo);
var D,C     : Double ;
    nbRec   : LongInt ;
    FRM     : TFRM ;
begin
  Q1.First ;
  while not Q1.EOF do
    begin
    // Delete ventilations analytiques + MAJ Sections
    DetruitLesSections(Q1, GeneTypeExo) ;

    // MAJ Comptes si écritures normales
    if E_QUALIFPIECE.Value='N' then
      begin
      D := Q1.FindField('E_DEBIT').AsFloat ;
      C := Q1.FindField('E_CREDIT').AsFloat ;

      // == Generaux ==
      // Préparation FRM
      Fillchar(FRM, SizeOf(FRM), #0) ;
      FRM.Cpt := Q1.FindField('E_GENERAL').AsString ;
      FRM.Deb := D ;
      FRM.Cre := C ;
      AttribParamsNew(FRM, D, C, GeneTypeExo) ;
      // MAJ section analytique
      nbRec := ExecReqINVNew(fbGene,FRM) ;
      if nbRec<>1 then V_PGI.IoError := oeSaisie ;

      // == Tiers ==
      if ( Q1.FindField('E_AUXILIAIRE').AsString <> '' ) then
        begin
        // Préparation FRM
        Fillchar(FRM,SizeOf(FRM),#0) ;
        FRM.Cpt := Q1.FindField('E_AUXILIAIRE').AsString ;
        FRM.Deb := D ;
        FRM.Cre := C ;
        AttribParamsNew(FRM, D, C, GeneTypeExo) ;
        // MAJ section analytique
        nbRec := ExecReqINVNew(fbAux, FRM) ;
        if nbRec <> 1
          then V_PGI.IoError := oeSaisie ;
        end ;

      // == Journaux ==
      // Préparation FRM
      Fillchar(FRM,SizeOf(FRM),#0) ;
      FRM.Cpt := Q1.FindField('E_JOURNAL').AsString ;
      FRM.Deb := D ;
      FRM.Cre := C ;
      AttribParamsNew(FRM, D, C, GeneTypeExo) ;
      // MAJ section analytique
      nbRec := ExecReqINVNew(fbJal, FRM) ;
      if nbRec <> 1
        then V_PGI.IoError := oeSaisie ;
      end  ;

    Q1.Next ;

    end ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : BOUSSERT Stéphane
Créé le ...... : 24/02/2003
Modifié le ... :   /  /
Description .. : MAJ des cumuls pour les sections analytiques
Suite ........ : +
Suite ........ : Retoune le nombre de lignes analytiques cibles
Mots clefs ... : 
*****************************************************************}
function TOF_CPMULMVT.InverseSoldesSections(Q1: TQuery; GeneTypeExo: TTypeExo): longint;
var nbLg, nbLa  : Longint ;
    D,C         : Double ;
    FRM         : TFRM ;
begin
  Result := 0 ;
  nbLg     := 0 ;
  if E_QUALIFPIECE.Value <> 'N' then Exit ;

  Q1.First ;

  // Parcours des lignes générales
  while not Q1.EOF do
    begin
    Inc(nbLg) ;
    D := Q1.FindField('Y_DEBIT').AsFloat ;
    C := Q1.FindField('Y_CREDIT').AsFloat ;
    // Préparation FRM
    Fillchar(FRM,SizeOf(FRM),#0) ;
    FRM.Cpt := Q1.FindField('Y_SECTION').AsString ; ;
    FRM.Deb := D ;
    FRM.Cre := C ;
    FRM.Axe := Q1.FindField('Y_AXE').AsString ;
    AttribParamsNew(FRM, D, C, GeneTypeExo) ;
    // MAJ section analytique
    nbLa := ExecReqINVNew(fbSect, FRM) ;
    if nbLa <> 1
      then V_PGI.IoError:=oeSaisie ;
    // Suivant
    Q1.Next ;
    end ;

  // Retour = nb de lignes traitées
  Result := nbLg ;

end;

function TOF_CPMULMVT.JePeuxDetruire(Q1: TQuery): boolean;
var
  St,CodeImmo : String ;
  {$IFDEF COMPTA}
  Clef : TClefPiece;
  {$ENDIF COMPTA}

begin
  Result := False ;
  Q1.First ;
  while not Q1.EOF do
    begin
    if Q1.FindField('E_VALIDE').AsString='X' then Exit ;
    if Q1.FindField('E_REFPOINTAGE').AsString<>'' then Exit ;
    if Q1.FindField('E_LETTRAGE').AsString<>'' then Exit ;
    if Q1.FindField('E_ETATLETTRAGE').AsString='PL' then Exit ;
    if Q1.FindField('E_ETATLETTRAGE').AsString='TL' then Exit ;
    if Q1.FindField('E_DATECOMPTABLE').AsDateTime<VH^.Encours.Deb then Exit ;
    if Q1.FindField('E_ECRANOUVEAU').AsString<>'N' then Exit ;
    if Q1.FindField('E_TIERSPAYEUR').AsString<>'' then
      begin
      St := Q1.FindField('E_PIECETP').AsString ;

    {$IFDEF COMPTA}
    Clef := ClefFromPieceTP(St);
    {Pour la nature de pièce, que ce soit celle de l'OD ou de la facture, ce n'est pas un critère
     déterminant : sur ces deux natures on gère les BAP
     Pour le QualifPiece, je pars du principe que les deux sont équivalentes}
    if not CanDeleteBap(Q1.FindField('E_NATUREPIECE').AsString,
                        Q1.FindField('E_QUALIFPIECE').AsString,
                        IntToStr(Clef.E_NUMEROPIECE),
                        UsDateTime(Clef.E_DATECOMPTABLE),
                        Clef.E_EXERCICE,
                        Clef.E_JOURNAL) then Exit;
    {$ENDIF COMPTA}

      if ((ListeTP.IndexOf(St)<0) and (St<>''))
        then ListeTP.Add(St) ;
      end ;
    CodeImmo := Q1.FindField('E_IMMO').AsString ;
    if CodeImmo<>'' then
      if Presence('IMMO','I_IMMO',CodeImmo) then Exit ;

    {$IFDEF COMPTA}
    if not CanDeleteBap(Q1.FindField('E_NATUREPIECE').AsString,
                        Q1.FindField('E_QUALIFPIECE').AsString,
                        Q1.FindField('E_NUMEROPIECE').AsString,
                        UsDateTime(Q1.FindField('E_DATECOMPTABLE').AsDateTime),
                        Q1.FindField('E_EXERCICE').AsString,
                        Q1.FindField('E_JOURNAL').AsString) then Exit;
    {$ENDIF COMPTA}

    Q1.Next ;
    end ;
  Result := True ;
end;

function TOF_CPMULMVT.UpdateSet(NumL: integer; EcartChange: boolean): String;
Var StM,StL : String ;
EtatL       : string;
begin
  EtatL := '';
  {$IFDEF COMPTA}
  if (Action='taSuppr') then  // ajout me fiche 19603
  begin
        if not AttenteLettrable then
          EtatL :=  ' E_ETATLETTRAGE="RI", E_ENCAISSEMENT="RIE", E_MODEPAIE="", E_DATEECHEANCE="'+stDate1900+'", '
        else
          EtatL :=  ' E_ETATLETTRAGE="AL", ';
  end;
  {$endif}
  StL := Copy( MessageListeSuppr[6] + ' ' + DateToStr(Date) + ' ' +
               MessageListeSuppr[7] + ' ' + V_PGI.UserName,           1,   35) ;
  if NumL=1
    then StM := '1'
    else StM := '-1' ;
  if EcartChange then
    begin
    Result := ' E_GENERAL="' + VH^.Cpta[fbGene].Attente + '", E_AUXILIAIRE="", E_DEBIT=' + StM + ','
//            + ' E_CREDIT=0, E_DEBITDEV=0, E_CREDITDEV=0, E_DEBITEURO=0, E_CREDITEURO=0, E_VALIDE="X", E_LIBELLE="' + StL + '",'
//            + ' E_COUVERTURE=0, E_COUVERTUREDEV=0, E_COUVERTUREEURO=0, E_LETTRAGE="", E_LETTRAGEDEV="-", E_TAUXDEV=1,'
            + ' E_CREDIT=0, E_DEBITDEV=0, E_CREDITDEV=0, E_VALIDE="X", E_LIBELLE="' + StL + '",'
            + ' E_COUVERTURE=0, E_COUVERTUREDEV=0, E_LETTRAGE="", E_LETTRAGEDEV="-", E_TAUXDEV=1,'
            + ' E_QTE1=0, E_QTE2=0, E_ECHE="-", E_ANA="-", E_NUMECHE=0, E_TYPEMVT="DIV",'
            + ' E_TVAENCAISSEMENT="-", '
            + ' E_CONTREPARTIEGEN="", E_CONTREPARTIEAUX="", E_CREERPAR="DET", '
            + ' E_DATEMODIF="' + UsTime(NowFutur) + '"' ;
   end
  else
   begin
   Result := ' E_GENERAL="' + VH^.Cpta[fbGene].Attente + '", E_AUXILIAIRE="", E_DEBIT=' + StM + ','
//           + ' E_CREDIT=0, E_DEBITDEV=' + StM + ', E_CREDITDEV=0, E_DEBITEURO=0, E_CREDITEURO=0, E_VALIDE="X", E_LIBELLE="' + StL + '",'
  //         + ' E_COUVERTURE=0, E_COUVERTUREDEV=0, E_COUVERTUREEURO=0, E_LETTRAGE="", E_LETTRAGEDEV="-", E_TAUXDEV=1,'
           + ' E_CREDIT=0, E_DEBITDEV=' + StM + ', E_CREDITDEV=0, E_VALIDE="X", E_LIBELLE="' + StL + '",'
           + ' E_COUVERTURE=0, E_COUVERTUREDEV=0, E_LETTRAGE="", E_LETTRAGEDEV="-", E_TAUXDEV=1,'
           + ' E_QTE1=0, E_QTE2=0, E_ECHE="-", E_ANA="-", E_NUMECHE=0, E_TYPEMVT="DIV",'
           + ' E_DEVISE="' + V_PGI.DevisePivot + '", E_TVAENCAISSEMENT="-", '
           + ' E_CONTREPARTIEGEN="", E_CONTREPARTIEAUX="", E_CREERPAR="DET", ' + EtatL
           + ' E_DATEMODIF="' + UsTime(NowFutur) + '"' ;
   end ;
end;

procedure TOF_CPMULMVT.InitFormModification;
begin
  // Init interface
  if Simul='N'
    then SetControlVisible('E_VALIDE',True)
    else SetControlVisible('E_VALIDE',False) ;

  E_QUALIFPIECE.Value := Simul ;
  if Action = 'taModif'
    then Ecran.Caption := MessageListe[1]
  else if Action = 'taConsult'
    then Ecran.Caption := MessageListe[0] ;

  InitCriteres ;
  // Paramètrage tables libres
  LibellesTableLibre(TTabSheet(GetControl('PZLibre', True)),'TE_TABLE','E_TABLE','E') ;

  {$IFDEF CONSOCERIC}
  E_QUALIFPIECE.Enabled := TRUE ;
  {$ENDIF}

  Simul := E_QUALIFPIECE.Value ;
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
    begin
      E_EXERCICE.Value      := VH^.Entree.Code ;
      E_DATECOMPTABLE.Text  := DateToStr(V_PGI.DateEntree) ;
      E_DATECOMPTABLE_.Text := DateToStr(V_PGI.DateEntree) ;
    end ;

  E_DATECREATION.Text   := stDate1900 ;
  E_DATECREATION_.Text  := stDate2099 ;
  E_DATEECHEANCE.Text   := stDate1900 ;
  E_DATEECHEANCE_.Text  := stDate2099 ;
  E_QUALIFPIECE.value   := Simul ;
  AttribTitre ;
// GP le 20/08/2008 21496 cf PositionneEtabUser + bas
//  PositionneEtabUser( E_ETABLISSEMENT ) ;
  if ANouveau then
  begin
    E_JOURNAL.DataType := 'ttJalANouveau' ;
    XX_WHERE.Text :=  'E_ECRANOUVEAU="H" OR E_ECRANOUVEAU="OAN"' ;
    E_NATUREPIECE.Value := 'OD' ;
    E_NATUREPIECE.Enabled := False ;
  end
  else
  begin
    if FBoSaisieParam and (Simul='N') then
      begin
      E_JOURNAL.DataType := 'CPJALSAISIEPARAM' ;
      XX_WHERE.Text      := 'E_ECRANOUVEAU="N"' ;
//      XX_WHEREBOR.Text   := '' ;
      end
    else
      begin
      E_JOURNAL.DataType := 'ttJalNonANouveau' ;
      XX_WHERE.Text := 'E_ECRANOUVEAU="N"' ;
      end ;
  end ;
  // GP le 20/08/2008 21496
  RetoucheHVCPoursaisie(E_JOURNAL) ;
  RetoucheHVCPoursaisie(E_ETABLISSEMENT) ;
  PositionneEtabUser( E_ETABLISSEMENT ) ;

{b fb 02/05/2006}
  if Engagement then begin
    if UpperCase(Action) = 'TAMODIF'   then
      XX_WHERE.Text := ' (CEN_STATUTENG="E" OR CEN_STATUTENG="P") '
    else if UpperCase(Action) = 'TACONSULT' then
      XX_WHERE.Text := '';
    XX_WHERE.Text := XX_WHERE.Text + ' AND E_CREERPAR="SAI" ' +
                     ' AND E_REFGESCOM="" ';
  end;

  {JP 26/06/07 : FQ TRESO 10491 : on vérouille les flux originaires de la Tréso
   JP 18/09/07 : FQ 21302 : on autorise la consultation des écritures de Tréso}
  if Action <> 'taConsult' then
    XX_WHEREGC.Text   := {'(E_REFGESCOM = "" OR E_REFGESCOM IS NULL) AND ' + A FAIRE ???}
                         '(E_QUALIFORIGINE <> "' + QUALIFTRESO + '" OR E_QUALIFORIGINE IS NULL)';


  if Prevision then begin
    XX_WHERE.Text := ' not exists (select * from cengagement ' +
		' where e_exercice=cen_exercice ' +
		' and e_journal=cen_journal ' +
		' and e_datecomptable=cen_datecomptable ' +
		' and e_numeropiece=cen_numeropiece ' +
		' and e_numligne=cen_numligne ' +
		' and e_qualifpiece=cen_qualifpiece ' +
		' and e_numeche=cen_numeche) ';
    end;
{e fb 02/05/2006}
  // Initialisation des critères
  if FbCharge then
  begin
    E_DATECOMPTABLE.Text  := DateToStr( DebutDeMois( FParams.DateC ) )  ;
    E_DATECOMPTABLE_.Text := DateToStr( FinDeMois( FParams.DateC ) ) ;
    E_JOURNAL.DataType    := 'TTJALFOLIO' ;
    E_JOURNAL.ItemIndex   := E_JOURNAL.Values.IndexOf( FParams.Jal ) ;
  end ;

{$IFDEF EAGLCLIENT}
  //TheMulQ.Manuel := FALSE ;
  //TheMUlQ.UpdateCriteres ;
{$ELSE}

  //Manuel := FALSE ;
  //TheMUlQ.UpdateCriteres ;
{$ENDIF}
  //SetControlText('XX_WHEREVIDE','') ;
end;

procedure TOF_CPMULMVT.InitFormSuppression;
begin
  // Activation de la multiselection
{$IFDEF EAGLCLIENT}
  FListe.MultiSelect := True ;
{$ELSE}
  FListe.MultiSelection := True ;
{$ENDIF}

  // Init variables
  ListePieces := TList.Create ;
  ListeTP     := HTStringList.Create ;

  // titre
  AttribTitreSuppr ;

  // Dates
  E_DATECREATION.Text  := StDate1900 ;
  E_DATECREATION_.Text := StDate2099 ;

  // ===> Limitation mode suppression :
  // combo Journal
  VirerChoixTous(E_JOURNAL) ;
  E_JOURNAL.Vide := False ;
  //SG6 25/01/05 FQ 15259
  E_JOURNAL.Datatype  := 'TTJALCRIT' ;
  if QueBOR then
    begin
    E_JOURNAL.PLUS    := ' AND J_NATUREJAL<>"ANO" AND (J_MODESAISIE="BOR" OR J_MODESAISIE="LIB")' ;
    XX_WHEREBOR.Text  := 'E_MODESAISIE="BOR" OR E_MODESAISIE="LIB"' ;
    end
  else
    begin
        if Simul = 'S' then // fiche 10064 trfs5 ajout me pour la suppression des écritures de simulations venant de sisco
        begin
             E_JOURNAL.PLUS   := ' AND J_NATUREJAL<>"ANO" ' ;
             XX_WHEREBOR.Text := 'E_QUALIFPIECE="S"' ;
        end
        else
        begin
             E_JOURNAL.PLUS   := ' AND J_NATUREJAL<>"ANO" AND (J_MODESAISIE="" OR J_MODESAISIE="-")' ;
             XX_WHEREBOR.Text := 'E_MODESAISIE="" OR E_MODESAISIE="-"' ;
        end;
    end ;

  // combo E_Etablissement
  PositionneEtabUser(E_ETABLISSEMENT) ;

  // combo Exercixe
  VirerChoixTous(E_EXERCICE) ;
  ChargeComboExo ;
  if VH^.Suivant.Code<>''
    then E_EXERCICE.Value := VH^.Suivant.Code
    else E_EXERCICE.Value := VH^.Encours.Code ;

  // Afficher Uniquement la 1ère ligne
  Rune.Checked        := True ;
  Rune.Enabled        := False ;

  if E_JOURNAL.Items.Count > 0
    then E_JOURNAL.ItemIndex := 0 ;

  // Condition supp
  XX_WHEREGC.Text := '(E_REFGESCOM = "" OR E_REFGESCOM IS NULL)';
  {JP 26/06/07 : FQ TRESO 10491 : on vérouille les flux originaires de la Tréso
   JP 18/09/07 : FQ 21302 : on autorise la consultation des écritures de Tréso}
  if Action <> 'taConsult' then
    XX_WHEREGC.Text := XX_WHEREGC.Text + ' AND (E_QUALIFORIGINE <> "' + QUALIFTRESO + '" OR E_QUALIFORIGINE IS NULL)';

  XX_Where.Text     := ' E_ECRANOUVEAU = "N" AND E_TRESOLETTRE <> "X" '
                     + 'AND (E_ETATLETTRAGE<>"TL" AND E_ETATLETTRAGE<>"PL") ' ;

  OkVuPiecesDetruites := False ;

{b fb 02/05/2006}
  if Engagement then begin
    if UpperCase(Action) = 'TASUPPR' then
      XX_WHERE.Text := ' CEN_STATUTENG="E" ';
    XX_WHERE.Text := XX_WHERE.Text + ' AND E_CREERPAR="SAI" ' +
                     ' AND E_REFGESCOM="" ';
  end;

  if Prevision then begin
    XX_WHERE.Text := ' not exists (select * from cengagement ' +
		' where e_exercice=cen_exercice ' +
		' and e_journal=cen_journal ' +
		' and e_datecomptable=cen_datecomptable ' +
		' and e_numeropiece=cen_numeropiece ' +
		' and e_numligne=cen_numligne ' +
		' and e_qualifpiece=cen_qualifpiece ' +
		' and e_numeche=cen_numeche) ';
    end;
{e fb 02/05/2006}
end;

procedure TOF_CPMULMVT.AttribTitreSuppr;
var Ind     : integer ;
    stValue : string;
begin
  Ind := 0 ;

  // Help Context
  stValue := E_QUALIFPIECE.Value ;
  if stValue = 'N' then
    begin
    Ind := 0  ;
    if QueBOR then
      begin
      Ind := 13;
      FEcran.HelpContext := 7263000
      end
    else FEcran.HelpContext := 7262000 ;
    end
  else if stValue = 'S' then
    begin
    Ind:=1  ;
    FEcran.HelpContext := 7283000 ;
    end
  else if stValue = 'U' then
    begin
    Ind:=2  ;
    FEcran.HelpContext := 7304000 ;
    end
  else if stValue = 'P' then
    begin
    Ind:=3  ;
    FEcran.HelpContext := 7319000 ;
    end
  else if stValue = 'R' then
    begin
    Ind:=10 ;
    FEcran.HelpContext := 7664000 ;
    end
  else if stValue = 'I' then  // Modif IFRS
    begin
    Ind:=13 ;
    FEcran.HelpContext := 7249400 ;
    end ;
{b fb 02/05/2006}
  if Engagement then begin
    Ind:=15;
    end;
{e fb 02/05/2006}
  // Titre
  FEcran.Caption := MessageListeSuppr[Ind] ;
  UpdateCaption(FEcran) ;

end;

procedure TOF_CPMULMVT.ValideClick;
{$IFDEF COMPTA}
var AA       : TActionFiche;
    lDossier : String ;
    lBoLocal : Boolean ;
    lMulQ    : TQuery ;

{$ENDIF}
begin
{$IFDEF COMPTA}
  if GetDataSet.Bof and GetDataSet.Eof then Exit ;

  if Action = 'taModif' then
    AA := taModif
  else
    AA := taConsult;

  // Réaffectation systématique de TheMulQ car vaut nil au 2ème passage
{$IFDEF EAGLCLIENT}
  lMulQ := FEcran.Q.TQ;
  lMulQ.Seek( FListe.Row - 1 );
{$ELSE}
  lMulQ := FEcran.Q;
{$ENDIF}

  lDossier := V_PGI.SchemaName ;
  lBoLocal := True ;

  if ( GetControlText('MULTIDOSSIER')<>'' ) then// EstMultiSoc and ...
    begin
    lDossier := GetField('SYSDOSSIER') ;
    // Consultation multi-Soc
    if lDossier <> V_PGI.SchemaName then
      begin
      lBoLocal := False ;
      AA       := taConsult ;
      end ;
    end ;

  if FBoSaisieParam or ( not lBoLocal ) then  // Saisie param ou consultation multi-Soc
    begin
    {b fb 02/05/2006}
      if Engagement then begin
        TrouveEtLanceSaisieParam( lMulQ, AA, 'p', FALSE, lDossier );
        FEcran.BChercheClick(nil);
       end else
    {e fb 02/05/2006}
    TrouveEtLanceSaisieParam( lMulQ, AA, GetControlText('E_QUALIFPIECE'), FALSE, lDossier )
    end
  else if ((sMode<>'-') and (sMode<>''))
    then LanceSaisieFolio(lMulQ,AA)
    else TrouveEtLanceSaisie(      lMulQ, AA, GetControlText('E_QUALIFPIECE') ) ;
  if (Action='taSuppr') then // ajout me fiche 19603
   AttenteLettrable :=  ExisteSQL ('SELECT G_LETTRABLE FROM GENERAUX WHERE G_GENERAL="' +VH^.Cpta[fbGene].Attente+'" and G_LETTRABLE="X"');
{$ENDIF}

end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 25/09/2006
Modifié le ... :   /  /    
Description .. : - 25/09/2006 - FB 18856 - on tient compte du mode de 
Suite ........ : synchro
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULMVT.ValideSupprClick;
var lBoDet : Boolean ;
begin
  if GetDataSet.Bof and GetDataSet.Eof then Exit ;

  // ajout me V590 pour interdire la suppression des écritures si liaison S1
  // fihce 19288
  if (Simul='N') then // FQ 18431 : ne pas faire le test que pour les écritures normales, les autres n'étant pas récupérées depuis S1
    if (GetParamsocSecur('SO_CPMODESYNCHRO',False) ) and ( ( GetParamSocSecur('SO_CPLIENGAMME' , '' ) = 'S1' ) or ( GetParamsocSecur('SO_CPLIENGAMME','') = 'S5' ))  then
      begin // CPLIENCOMPTABILITE
      PGiBox ('Le dossier est en liaison avec ' + RechDom('CPLIENCOMPTABILITE', GetParamSocSecur('SO_CPLIENGAMME' , '' ), false ) + ', la suppression est interdite.', 'Suppression des Ecritures');
      FListe.ClearSelected;
      exit;
      end;

  // FQ 18445 SBO ajout bouton tout sélectionné pour les travaux de simul
  if ((Simul='S') or (Simul='U')) and (FListe.AllSelected)
    then lBoDet := DetruitSelectionSimul
    else lBoDet := DetruitSelection ;
  if lBoDet
    then FEcran.BChercheClick(Nil) ;

  OkVuPiecesDetruites := False ;
end;

procedure TOF_CPMULMVT.BDetruitesClick(Sender: TObject);
begin
  if ListePieces.Count>0 then
    begin
    OkVuPiecesDetruites := True ;
    VisuPiecesGenere(ListePieces,EcrGen,0) ;
    end ;
end;

procedure TOF_CPMULMVT.OnClose;
begin

  if action = 'taSuppr' then
    begin
    // Visu pieces détruites
    if (ListePieces.Count>0) and (not OkVuPiecesDetruites) then
      if HShowMessage(MessageListeSuppr[8],'','') = mrYes
        then VisuPiecesGenere(ListePieces,EcrGen,0) ;
    // Libération TList
    VideListe(ListePieces) ;
    ListePieces.Free ;
    ListeTP.Clear ;
    ListeTP.Free ;

    _Bloqueur('nrAucun',False) ; // FQ 17425

    end ;

  inherited ;
end;

procedure TOF_CPMULMVT.VirerChoixTous(vCombo: THValComboBox);
begin
  if not vCombo.vide then Exit ;
  vCombo.Items.Delete(0) ;
  vCombo.Values.Delete(0) ;
end;

procedure TOF_CPMULMVT.BZoomPieceClick(Sender: TObject);
begin
  ValideClick ;
end;

procedure TOF_CPMULMVT.BValideClick(Sender: TObject);
begin
  if Action = 'taSuppr'
   then ValideSupprClick
   else ValideClick ;
end;

// 14386
procedure TOF_CPMULMVT.AfterShow;
begin
  inherited;
  FEcran.BChercheClick(Nil) ;
end;

{JP 06/10/05 : Gestion des filtres et des listes avec la nouvelle gestion de l'agl => il n'est
               plus possible de désolidariser les filtres et les listes, j'ai donc créé autant
               de listes qu'il y a de filtres
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULMVT.AttribListe;
{---------------------------------------------------------------------------------------}
var
  LaListe : string;
begin
  {Premier Filtre, le type d'action}
       if UpperCase(Action) = 'TAMODIF'   then LaListe := 'CPMODECR'
  else if UpperCase(Action) = 'TACONSULT' then LaListe := 'CPVISECR'
  else if UpperCase(Action) = 'TASUPPR'   then LaListe := 'CPSUPECR';

  {second filtre, le type d'écritures}
{b fb 02/05/2006}
  if Engagement then
    LaListe := LaListe + 'E'
{e fb 02/05/2006}
  else 
       if ANouveau       then LaListe := LaListe + 'H'
  // SBO 15/05/2007 : FQ 20311 : détails des listes pour chaque entrée par E_QUALIFPIECE
  else if FBoSaisieParam then
       begin
       LaListe := LaListe + 'M' ;
       if Simul<>'N' then
         LaListe := LaListe + Simul ;
       end
  else LaListe := LaListe + Simul;

  {Mise à jour de la liste, du filtre et du ParamDBG}
  TFMul(FEcran).SetDBListe(LaListe);
end;

{$IFDEF COMPTA}
{---------------------------------------------------------------------------------------}
function TOF_CPMULMVT.CanDeleteBap(Nat, Qual, NumPiece, UsDateCpt, Exo, Jal : string) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := True;
  if Qual <> 'N' then Exit;
  if not FBoGestionBap then Exit;
  if not IsJournalBap(Nat, Jal, '') then Exit; 
  if not ExisteSQL('SELECT BAP_JOURNAL FROM CPBONSAPAYER WHERE BAP_STATUTBAP = "' + sbap_Definitif + '" AND ' +
                   'BAP_JOURNAL = "' + Jal + '" AND BAP_EXERCICE = "' + Exo + '" AND ' +
                   'BAP_DATECOMPTABLE = "' + UsDateCpt + '" AND BAP_NUMEROPIECE = ' + NumPiece) then Exit;
  Result := False;
end;
{$ENDIF COMPTA}

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 02/06/2006
Modifié le ... :   /  /    
Description .. : - FB 10678 - LG - 02/06/2006 - gestion des devises ds les 
Suite ........ : ano dynamiques
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULMVT.DetruitLesAnoDyna( vTOB : TOB ) ;
var
 i : integer ;
 lTSA : TList ;
 lZCompte : TZCompte ;
 lStCompte : string ;
 lStNat : string ;
begin

 lZCompte := TZCompte.Create ;
 lTSA     := TList.Create ;

 for i := 0 to vTOB.Detail.Count - 1 do
  begin
   lStCompte := vTOB.Detail[i].GetValue('E_GENERAL') ;
   if lZCompte.GetCompte(lStCompte) > - 1 then
    lStNat := lZCompte.GetValue('G_NATUREGENE') ;
   AjouteAno(lTSA,vTOB.Detail[i],lStNat,true) ;
  end ;

  if not ExecReqMAJAno(lTSA) then
   begin
    V_PGI.IoError := oeSaisie ;
    exit ;
   end ;

 lZCompte.Free ;
 for i:=0 to lTSA.Count-1 do
  if assigned(lTSA[i]) then Dispose(lTSA[i]);
 lTSA.Free ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 25/09/2006
Modifié le ... :   /  /    
Description .. : // FQ18445 : SBO 25/09/2006
Suite ........ : Suppression en masse des écritures sur tout sélectionner
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULMVT.DetruitAllSimul;
begin
  // Ajustement pour récupérer toutes les lignes des pièces
  SetControlText('E_NUMECHE',	'9999') ;
  SetControlText('E_NUMLIGNE',	'0') ;
  SetControlText('E_NUMLIGNE_',	'9999') ;

   ExecuteSQL('DELETE FROM ECRITURE ' + RecupWhereCritere(TPageControl(GetControl('Pages', True))) ) ;

  // Ajustement pour récupérer toutes les lignes des pièces
  SetControlText('E_NUMECHE',	'1') ;
  SetControlText('E_NUMLIGNE',	'1') ;
  SetControlText('E_NUMLIGNE_',	'1') ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 22/09/2006
Modifié le ... : 25/09/2006
Description .. : // FQ18445 : 25/09/2006
Suite ........ : Retourne Vrai si la sélection peutêtre détruite dans le cadre
Suite ........ : des écritures de simulation/situation en mode "tout 
Suite ........ : sélectionner"
Mots clefs ... : 
*****************************************************************}
function TOF_CPMULMVT.JePeuxDetruireAllSimul: Boolean;
begin

  // Ajustement pour récupérer toutes les lignes des pièces
  SetControlText('E_NUMECHE',	'9999') ;
  SetControlText('E_NUMLIGNE',	'0') ;
  SetControlText('E_NUMLIGNE_',	'9999') ;

  // TEST lignes validée ou sur exo clos
  result := not ExisteSQL('SELECT E_JOURNAL FROM ECRITURE ' + RecupWhereCritere( TPageControl(GetControl('Pages', True)))
                                  + ' AND (E_VALIDE="X" OR E_DATECOMPTABLE<"' + UsDateTime(VH^.Encours.Deb) + '" ) ' ) ;

  // Ajustement pour récupérer toutes les lignes des pièces
  SetControlText('E_NUMECHE',	'1') ;
  SetControlText('E_NUMLIGNE',	'1') ;
  SetControlText('E_NUMLIGNE_',	'1') ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULMVT.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 25/09/2006
Modifié le ... :   /  /
Description .. : // FQ18445 : SBO 25/09/2006
Suite ........ : Gestion de la suppression par action "tout sélectionner"
Mots clefs ... :
*****************************************************************}
function TOF_CPMULMVT.DetruitSelectionSimul : Boolean ;
var lInQte  : integer ;
    lQCount : TQuery ;
begin
  Result  := False ;

  // FQ 18445 suppression sur tout sélectionner
  if not (FListe.AllSelected) then Exit ;

  // comptage et demande de confirmation
  lInQte  := 0 ;
  lQCount := OpenSql('SELECT COUNT(DISTINCT E_NUMEROPIECE) QTE FROM ECRITURE ' + RecupWhereCritere( TPageControl(GetControl('Pages', True))), True ) ;
  if not lQCount.Eof
    then lInQte := lQCount.FindField('QTE').AsInteger ;
  Ferme(lQCount) ;
  // message
  if PgiAsk('Votre sélection comporte ' + IntToStr(lInQte) + ' pièces comptables. Confirmez-vous leur suppression ?', Ecran.Caption ) <> mrYes then
    Exit ;

  // Test si suppression possible
  if JePeuxDetruireAllSimul then
    begin
    // Suppression
    if Transactions(DetruitAllSimul, 1)<>oeOK then
      begin
      MessageAlerte( MessageListeSuppr[9] ) ;
      exit ;
      end ;
    end
  else
    begin
    PgiInfo( MessageListeSuppr[16] ) ;
    exit ;
    end ;

  Result  := True ;

end;


Initialization
  registerclasses ( [ TOF_CPMULMVT ] ) ;
end.

