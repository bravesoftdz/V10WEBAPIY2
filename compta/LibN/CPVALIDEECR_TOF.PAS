{***********UNITE*************************************************
Auteur  ...... : YMO
Créé le ...... : 03/10/2006
Modifié le ... :   /  /
Description .. : Source TOF des FICHES : CPVALIDEECR et CPDEVALIDEECR
Mots clefs ... : TOF;VALIDEECR
*****************************************************************}
Unit CPVALIDEECR_TOF                                               ;

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
     AGLInit,               // TheMulQ
     Zcompte,
     SaisComm,
     dialogs,              // Pour les procédures de MAJ des comptes
     HSysMenu,
     Windows,              // VK_F10
     ULibExercice;         // CQuelExercice

procedure CPLanceFiche_ValideEcriture(vStRange, vStLequel, vStArgs : string);

Type
  TOF_CPVALIDEECR = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose ; override ;
  private
    FEcran    : TFMul ;
    Action    : String ;    // 'D'évalidation ou 'V'alidation
    ANouveau  : Boolean ;
    FbCharge  : Boolean ;
    sMode     : String ;
    MsgBox    : THMsgBox;
    FListeEcr : TOB;
    {$IFNDEF EAGLCLIENT}
    HMTrad: THSystemMenu;
    {$ENDIF}
    // Eléments interface
    E_JOURNAL        : THValComboBox ;
    E_NATUREPIECE    : THValComboBox ;
    E_DEVISE         : THValComboBox ;
//    E_QUALIFPIECE    : THValComboBox ;
    E_EXERCICE       : THValComboBox ;
    E_ETABLISSEMENT  : THValComboBox ;
    E_DATECOMPTABLE  : THEdit ;
    E_DATECOMPTABLE_ : THEdit ;
    E_DATECREATION   : THEdit ;
    E_DATECREATION_  : THEdit ;
    E_DATEECHEANCE   : THEdit ;
    E_DATEECHEANCE_  : THEdit ;
    XX_WHERE         : THEdit ;
  {$IFDEF EAGLCLIENT}
    FListe : THGrid ;
  {$ELSE}
    FListe : THDBGrid ;
  {$ENDIF}

    // Evènements fiche
    procedure E_EXERCICEChange(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure OuvrirClick ;
    procedure ValideClick ;
    procedure BOuvrirClick ( Sender : TObject ) ;
    procedure BValideClick ( Sender : TObject ) ;
    procedure BChercheClick(Sender: TObject);
// Autres procédures
    procedure InitCriteres ;
    procedure AttribTitre ;
    procedure AttribListe;
    Procedure InitFormModification ;
    procedure InitMsgBox;
    procedure FormKeyDown( Sender : TObject; var Key: Word; Shift: TShiftState );
    procedure AuxiElipsisClick(Sender : TObject);
(*  BVE 24.08.07 : Plus utilisé
    procedure OnChercheClick(Sender: TOBject);
    procedure ChargeLesEcritures;
    function  ConstruitRequete : String ; *)
  end ;

Implementation

uses                               
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  UlibEcriture,
  {$IFDEF eAGLCLIENT}
  MenuOLX
  {$ELSE}
  MenuOLG
  {$ENDIF eAGLCLIENT}
  , Constantes
  , UTofMulParamGen; {26/04/07 YMO F5 sur Auxiliaire }

Const MessageListe : Array[0..1] of String =( 'Validation des écritures','Dévalidation des écritures');
type ttCarAlpha = set of char;
var StopCherche : boolean;

//==============================================================================
procedure CPLanceFiche_ValideEcriture(vStRange, vStLequel, vStArgs : string);
begin

   AGLLanceFiche ('CP','CPVALIDEECR', vStRange, vStLequel, vStArgs );

end;


//==============================================================================

procedure TOF_CPVALIDEECR.OnArgument (S : String ) ;
var
  wh : string;
begin

  // Récupération des arguments
  if Trim(S) <> '' then
    Action 	:= ReadTokenSt(S) ;

  FbCharge := FALSE ; // ??
  sMode	   := '-' ;   // ??

  If Action = 'D' Then
      SetControlProperty('E_VALIDE','State',cbChecked)
  else
      SetControlProperty('E_VALIDE','State',cbUnChecked);

  // Récup interface
  FEcran := TFMul(Ecran) ;
{$IFDEF EAGLCLIENT}
  FListe  := THGrid( GetControl('FListe',True) ) ;
  TheMulQ := FEcran.Q.TQ;
{$ELSE}
  FListe  := THDBGrid( GetControl('FListe',True)) ;
  TheMulQ := FEcran.Q;
{$ENDIF}

  FListeEcr := TOB.Create('', nil, -1);
  GetControl('E_VALIDE',True).Enabled:=False;

  E_JOURNAL        := THValComboBox(GetControl('E_JOURNAL', True)) ;
  E_NATUREPIECE    := THValComboBox(GetControl('E_NATUREPIECE', True)) ;
  E_DEVISE         := THValComboBox(GetControl('E_DEVISE', True)) ;
//  E_QUALIFPIECE    := THValComboBox(GetControl('E_QUALIFPIECE', True)) ;
  E_EXERCICE       := THValComboBox(GetControl('E_EXERCICE', True)) ;
  E_ETABLISSEMENT  := THValComboBox(GetControl('E_ETABLISSEMENT', True)) ;
  E_DATECOMPTABLE  := THEdit(GetControl('E_DATECOMPTABLE', True))  ;
  E_DATECOMPTABLE_ := THEdit(GetControl('E_DATECOMPTABLE_', True)) ;
  E_DATECREATION   := THEdit(GetControl('E_DATECREATION', True))  ;
  E_DATECREATION_  := THEdit(GetControl('E_DATECREATION_', True)) ;
  E_DATEECHEANCE   := THEdit(GetControl('E_DATEECHEANCE', True)) ;
  E_DATEECHEANCE_  := THEdit(GetControl('E_DATEECHEANCE_', True)) ;
  XX_WHERE         := THEdit(GetControl('XX_WHERE', True)) ;

  Ecran.OnKeyDown  := FormKeyDown;

	// Réaffectation des évènements
  E_EXERCICE.OnChange  := E_EXERCICEChange ;
//  E_EXERCICE.Vide      :=True;
//  E_EXERCICE.VideString:='<<Tous>>';

  FListe.OnDblClick 					  := FListeDblClick ;

  TButton(GetControl('BOuvrir', True)).OnClick 		  := BOuvrirClick ;
  TButton(GetControl('BValider', True)).OnClick 		  := BValideClick ;

  // On affiche seulement la 1ère ligne des pièces
  SetControlText('E_NUMLIGNE',	'1') ;
  SetControlText('E_NUMLIGNE_',	'1') ;
  SetControlText('E_NUMECHE',		'1') ;

  TToolbarButton97( GetControl('BCHERCHE', True)   ).OnClick := BChercheClick ;
  // Puisque le 'Plus' des datatype n'est pas repris lorsque '<<Tous>>' est choisi sur E_EXERCICE par exemple, il faut reprendre les restrictions
  wh := ' AND E_CREERPAR<>"DET" AND (E_QUALIFPIECE="N" OR E_QUALIFPIECE="I")' +
        ' AND E_EXERCICE IN (SELECT EX_EXERCICE FROM EXERCICE WHERE EX_ETATCPTA="OUV" OR EX_ETATCPTA="CPR")';
  {JP 26/06/07 : Par contre on ne touche pas aux pièces de Tréso}
  wh := wh + ' AND (E_QUALIFORIGINE <> "' + QUALIFTRESO + '" OR E_QUALIFORIGINE = "" OR E_QUALIFORIGINE IS NULL)';
  SetControlText('XX_WHERE', wh);

  inherited ;

  InitFormModification;
  InitMsgBox;
  AttribListe;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
    THEdit(GetControl('E_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;

end ;

procedure TOF_CPVALIDEECR.InitCriteres;
begin
  if VH^.Precedent.Code<>''
    then E_DATECOMPTABLE.Text := DateToStr(VH^.Precedent.Deb)
    else E_DATECOMPTABLE.Text := DateToStr(VH^.Encours.Deb) ;
  E_DATECOMPTABLE_.Text := DateToStr(V_PGI.DateEntree) ;

   // Init des valeurs

  E_JOURNAL.Value  := '';
  E_NATUREPIECE.Value  := '';
  E_DEVISE.Value := '';
end;

procedure TOF_CPVALIDEECR.AttribTitre;
BEGIN

  // HelpContext
  // Ecran.HelpContext:=7727000 ;

  // Titre
  If Action='V' then Ecran.Caption := MessageListe[0] ;
  If Action='D' then Ecran.Caption := MessageListe[1] ;
  
  UpdateCaption(Ecran) ;
end;


procedure TOF_CPVALIDEECR.E_EXERCICEChange(Sender: TObject);
begin
 	ExoToDates( GetControlText('E_EXERCICE'), TEdit(GetControl('E_DATECOMPTABLE', True)), TEdit(GetControl('E_DATECOMPTABLE_', True)) ) ;
	if ((GetControlText('E_EXERCICE')='')) then
  	BEGIN
    SetControlText('E_DATECOMPTABLE',		stDate1900) ;
    SetControlText('E_DATECOMPTABLE_',	stDate2099) ;
    END ;
end;


procedure TOF_CPVALIDEECR.FListeDblClick(Sender: TObject);
begin
  OuvrirClick ;
end;


procedure TOF_CPVALIDEECR.InitFormModification;
begin
  // Init interface
  InitCriteres ;
  // Paramètrage tables libres
  LibellesTableLibre(TTabSheet(GetControl('PZLibre', True)),'TE_TABLE','E_TABLE','E') ;

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

  E_DATECREATION.Text   := stDate1900 ;
  E_DATECREATION_.Text  := stDate2099 ;
  E_DATEECHEANCE.Text   := stDate1900 ;
  E_DATEECHEANCE_.Text  := stDate2099 ;

  AttribTitre ;
  PositionneEtabUser( E_ETABLISSEMENT ) ;
  StopCherche := True; // Pas d'affichage dès l'ouverture

  E_JOURNAL.Plus:=' AND J_FERME="-" AND J_NATUREJAL<>"ANO" AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ODA" And J_NATUREJAL<>"ANA"';

  //positionnement sur le 1er journaél de la liste
  If (E_JOURNAL.Values.Count > 1) and (E_JOURNAL.Values[0]='') then
      E_JOURNAL.Value:=E_JOURNAL.Values[1];
end;


procedure TOF_CPVALIDEECR.OuvrirClick;
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

procedure TOF_CPVALIDEECR.ValideClick;
var lMulQ : TQuery ;
    i     : integer;
    Sql, StValide,
    SQLWhere, SQLBaseE, SQLBaseY : string;
    ListePiece : HTStringList ;
begin


If Action='D' then //YMO 28/11/2006 FQ19232
 PgiInfo('En référence au BOI 13 L-1-06 N° 12 du 24 janvier 2006 paragraphe 23 qui rappelle l''interdiction de ' + #13#10 +
         'toute modification et/ou suppression d''écriture validée, nous vous conseillons de « tracer » l’information ' + #13#10 +
         'par toute méthode à votre disposition (par ex le bloc note).', 'Dévalidation des écritures');


{$IFDEF EAGLCLIENT}
  lMulQ := FEcran.Q.TQ;
  lMulQ.Seek( FListe.Row - 1 );
{$ELSE}
  lMulQ := FEcran.Q;
{$ENDIF}

  ListePiece:=HTStringList.Create ;
  ListePiece.Clear ;

{$IFDEF CERTIFNF}
  If Action = 'D' then
     ListePiece.Add('ANNVALECR ')
  Else
     ListePiece.Add('VALECR ');
{$ELSE}    
  ListePiece.Sorted:=TRUE ;
{$ENDIF}

  If Action = 'D' Then StValide:='-' else StValide:='X' ;
  
  If TFmul(Ecran).FListe.AllSelected then
  begin
      LMulQ.First;
      While not LMulQ.EOF do
      begin    //YMO 30/11/2006 Valable pour les bordereaux et les pièces
          If LMulQ.BOF then SQLBaseE:=SQLBaseE+' AND ' else SQLBaseE:=SQLBaseE+' OR ';
          SQLBaseE:=SQLBaseE+'(E_NUMEROPIECE='+LMulQ.FindField('E_NUMEROPIECE').AsString
          +' AND E_JOURNAL="'+LMulQ.FindField('E_JOURNAL').AsString
          +'" AND E_EXERCICE="'+QuelExoDt(LMulQ.FindField('E_DATECOMPTABLE').AsDateTime)
          +'" AND E_PERIODE="'+IntToStr(GetPeriode(LMulQ.FindField('E_DATECOMPTABLE').AsDateTime))
          +'" )';

          If LMulQ.BOF then SQLBaseY:=SQLBaseY+' AND ' else SQLBaseY:=SQLBaseY+' OR ';
          SQLBaseY:=SQLBaseY+'(Y_NUMEROPIECE='+LMulQ.FindField('E_NUMEROPIECE').AsString
          +' AND Y_JOURNAL="'+LMulQ.FindField('E_JOURNAL').AsString
          +'" AND Y_EXERCICE="'+QuelExoDt(LMulQ.FindField('E_DATECOMPTABLE').AsDateTime)
          +'" AND Y_PERIODE="'+IntToStr(GetPeriode(LMulQ.FindField('E_DATECOMPTABLE').AsDateTime))
          +'" )';

          ListePiece.Add(LMulQ.FindField('E_JOURNAL').AsString+'-'
                        +IntToStr(GetPeriode(LMulQ.FindField('E_DATECOMPTABLE').AsDateTime))+'-'
                        +LMulQ.FindField('E_NUMEROPIECE').AsString);
          
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

    //YMO 30/11/2006 Valable pour les bordereaux et les pièces
    If i<>0 then SQLBaseE:=SQLBaseE+' OR ' else SQLBaseE:=SQLBaseE+' AND ';
    SQLBaseE:=SQLBaseE+'(E_NUMEROPIECE='+LMulQ.FindField('E_NUMEROPIECE').AsString
    +' AND E_JOURNAL="'+LMulQ.FindField('E_JOURNAL').AsString
    +'" AND E_EXERCICE="'+QuelExoDt(LMulQ.FindField('E_DATECOMPTABLE').AsDateTime)
    +'" AND E_PERIODE="'+IntToStr(GetPeriode(LMulQ.FindField('E_DATECOMPTABLE').AsDateTime))
    +'" )';

    If i<>0 then SQLBaseY:=SQLBaseY+' OR ' else SQLBaseY:=SQLBaseY+' AND ';
    SQLBaseY:=SQLBaseY+'(Y_NUMEROPIECE='+LMulQ.FindField('E_NUMEROPIECE').AsString
    +' AND Y_JOURNAL="'+LMulQ.FindField('E_JOURNAL').AsString
    +'" AND Y_EXERCICE="'+QuelExoDt(LMulQ.FindField('E_DATECOMPTABLE').AsDateTime)
    +'" AND Y_PERIODE="'+IntToStr(GetPeriode(LMulQ.FindField('E_DATECOMPTABLE').AsDateTime))
    +'" )';

    ListePiece.Add(LMulQ.FindField('E_JOURNAL').AsString+'-'
                  +IntToStr(GetPeriode(LMulQ.FindField('E_DATECOMPTABLE').AsDateTime))+'-'
                  +LMulQ.FindField('E_NUMEROPIECE').AsString);

  end;

  FiniMove;

  SQLWhere:=SQLBaseE;

  // Mode Pièce ou Bordereau ?
//  If ExisteSQL('SELECT J_MODESAISIE FROM JOURNAL WHERE J_MODESAISIE="-" AND J_JOURNAL="'+E_JOURNAL.Value+'"') then

  SQLWhere:=SQLWhere+' AND (E_QUALIFPIECE="N" OR E_QUALIFPIECE="I")';

  // FQ19230 28/11/2006 SQLWhere:=SQLWhere+' AND (E_NUMLIGNE >= 1 AND E_NUMLIGNE <= 1) ';

  // mise a jour des ecritures
  Sql := 'UPDATE ECRITURE SET E_VALIDE="'+StValide
  + '" WHERE E_ECRANOUVEAU="N"' + SQLWhere;
  ExecuteSQL(Sql);

  SQLWhere:=SQLBaseY;

  SQLWhere:=SQLWhere+' AND (Y_QUALIFPIECE="N" OR Y_QUALIFPIECE="I")';

  Sql := 'UPDATE ANALYTIQ SET Y_VALIDE="'+StValide
  + '" WHERE Y_TYPEANALYTIQUE="-" AND Y_ECRANOUVEAU="N"' + SQLWhere;

  ExecuteSQL(Sql);
         
  { BVE 29.08.07 : Mise en place d'un nouveau tracage  }
{$IFNDEF CERTIFNF}
  For i:=0 To ListePiece.Count-1 Do
  Begin
     If Action = 'D' Then
        CPEnregistreLog('ANNVALECR ' + ListePiece[i])
     Else
        CPEnregistreLog('VALECR ' + ListePiece[i]);
  End;
{$ELSE}
  CPEnregistreJalEvent('CVE','Validation d''écritures',ListePiece);
{$ENDIF}

  ListePiece.Free ;

end;

procedure TOF_CPVALIDEECR.OnClose;
begin
  FListeEcr.Free;
  inherited ;
end;


procedure TOF_CPVALIDEECR.BOuvrirClick(Sender: TObject);
begin
 OuvrirClick ;
end;

procedure TOF_CPVALIDEECR.BValideClick(Sender: TObject);
begin
 //06/12/2006 YMO Norme NF 203
{$IFDEF CERIFNF}
 if (Action = 'D') and (GetParamSocSecur('SO_CPCONFORMEBOI', False)) then
 begin
   PGIInfo('Pour la conformité stricte avec la norme NF 203 (et le BOI du 24/01/2006) cette fonction n''est plus disponible',Ecran.Caption);
   Exit;
 end;
{$ENDIF}
 If (Fliste.NbSelected=0) and (not Fliste.AllSelected) then Exit ;
 If ((Action='D') And (MsgBox.Execute(1,'','')=mrYes))
 Or ((Action<>'D') And (MsgBox.Execute(0,'','')=mrYes)) then
      ValideClick ;

 FListe.ClearSelected;
 TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVALIDEECR.AttribListe;
{---------------------------------------------------------------------------------------}
var
  LaListe : string;
begin
  If Action='D' then
    LaListe := 'CPDEVALIDEECR'
  else
    LaListe := 'CPVALIDEECR' ;

  {Mise à jour de la liste, du filtre et du ParamDBG}
  TFMul(FEcran).SetDBListe(LaListe);
end;


procedure TOF_CPVALIDEECR.InitMsgBox;
begin
  MsgBox := THMsgBox.create(FMenuG);
  MsgBox.Mess.Add('0;Validation;Ce traitement va valider toutes les pièces sélectionnées. Désirez-vous continuer?;Q;YN;Y;;');
  MsgBox.Mess.Add('1;Dévalidation;Ce traitement va dévalider toutes les pièces sélectionnées. Désirez-vous continuer?;Q;YN;Y;;');
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/10/2006
Modifié le ... :  /  /
Description .. : On n'affiche pas les pièces à l'entrée dans le MUL
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPVALIDEECR.BChercheClick(Sender: TObject);
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
  
{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 26/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVALIDEECR.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

(* BVE 24.08.07 Plus utilisé ?
procedure TOF_CPVALIDEECR.OnChercheClick(Sender: TOBject);
begin
   ChargeLesEcritures;
end;

procedure TOF_CPVALIDEECR.ChargeLesEcritures;
var
  St: string;
begin
//    FListe.SynEnabled := False;
    FListeEcr.ClearDetail;
//    FListe.VidePile(False);

    // SBO 14/10/2005 : Modif de la requête pour optimisation
    St := ConstruitRequete ;

    // chargement et affichage des ecritures
    FListeEcr.LoadDetailDBFromSQl('ECRITURE',St,false);

//    FListeEcr.PutGridDetail(THGrid(GetControl('FLISTE')),false,false,'E_NATUREPIECE;E_DATECOMPTABLE;E_NUMEROPIECE;E_GENERAL;E_AUXILIAIRE;E_REFINTERNE;E_LIBELLE;E_DEBIT;E_CREDIT;E_PERIODE;E_NUMLIGNE');
    FListeEcr.PutGridDetail(THGrid(GetControl('FLISTE')),false,false,'E_DATECOMPTABLE;E_NUMEROPIECE;E_GENERAL;E_AUXILIAIRE');
end;


function TOF_CPVALIDEECR.ConstruitRequete: String;
var St  : String ;
begin

  St := 'SELECT E1.E_JOURNAL, E1.E_DATECOMPTABLE, E1.E_NATUREPIECE, E1.E_GENERAL, E1.E_AUXILIAIRE,';
  St :=St + ' E1.E_REFINTERNE, E1.E_DEBIT, E1.E_CREDIT, E1.E_MODESAISIE, E1.E_EXERCICE, E1.E_NUMECHE,';
  St :=St + ' E1.E_QUALIFPIECE, E1.E_VALIDE, E1.E_NUMEROPIECE, E1.E_NUMLIGNE, E1.E_PERIODE';
  St :=St + ' FROM ECRITURE E1';
  St :=St + ' LEFT JOIN ECRITURE E2 ON E1.E_JOURNAL=E2.E_JOURNAL';
  St :=St + ' AND E1.E_EXERCICE=E2.E_EXERCICE';
  St :=St + ' AND E1.E_DATECOMPTABLE=E2.E_DATECOMPTABLE';
  St :=St + ' AND E1.E_NUMEROPIECE=E2.E_NUMEROPIECE';
  St :=St + ' AND E1.E_QUALIFPIECE=E2.E_QUALIFPIECE';
  St :=St + ' WHERE (E1.E_NUMECHE <= 1';
  St :=St + ' AND (AND E1.E_CREERPAR<>"DET"';
  St :=St + ' AND (E1.E_QUALIFPIECE="N" OR E1.E_QUALIFPIECE="I")';
  St :=St + ' AND E1.E_EXERCICE IN';
  St :=St + ' (SELECT EX_EXERCICE FROM EXERCICE';
  St :=St + ' WHERE EX_ETATCPTA="OUV" OR EX_ETATCPTA="CPR")';
  St :=St + ' AND (E1.E_QUALIFORIGINE <> "TRO" OR E1.E_QUALIFORIGINE = "" OR E1.E_QUALIFORIGINE IS NULL))';
  St :=St + ' AND E1.E_NUMLIGNE <= 1 AND E1.E_VALIDE<>"X")';
  St :=St + ' ORDER BY E1.E_DATECOMPTABLE DESC';

  {
  if (GetCheckBoxState('SELECTION_PIECE') = cbChecked)
    then Pf  := 'E1.'
    else Pf  := '' ;

  // **** Les Champs ****
  St := '@NOCONF@SELECT DISTINCT ' + Pf + 'E_NATUREPIECE, ' + Pf + 'E_DATECOMPTABLE, ' + Pf + 'E_NUMEROPIECE, '
                           + Pf + 'E_GENERAL, '     + Pf + 'E_AUXILIAIRE, '    + Pf + 'E_REFINTERNE, '
                           + Pf + 'E_LIBELLE, '     + Pf + 'E_DEBIT, '         + Pf + 'E_CREDIT, '
                           + Pf + 'E_PERIODE, '     + Pf + 'E_NUMLIGNE,'       + Pf + 'E_JOURNAL, '
                           + Pf + 'E_QUALIFPIECE, ' + Pf + 'E_EXERCICE,'       + Pf + 'E_NUMECHE, '
                           + Pf + 'E_NUMGROUPEECR ';  //FQ19187 YMO 12/03/07 Sélection en mode bordereau par groupe et non pas folio

  // *** Les tables ***
  St :=   St + 'FROM ECRITURE' ;
  if (GetCheckBoxState('SELECTION_PIECE') = cbChecked) then
    St := St  + ' E1 LEFT JOIN ECRITURE E2 ON E1.E_JOURNAL=E2.E_JOURNAL AND E1.E_EXERCICE=E2.E_EXERCICE'
                   + ' AND E1.E_DATECOMPTABLE=E2.E_DATECOMPTABLE AND E1.E_NUMEROPIECE=E2.E_NUMEROPIECE'
                   + ' AND E1.E_QUALIFPIECE=E2.E_QUALIFPIECE' ;

  // **** Les Conditions ****
  St := St  + ' WHERE ' + GetWhereSQL ;

  // *** Tri ***
  St := St + ' ORDER BY ' + Pf + 'E_PERIODE, ' + Pf + 'E_NUMEROPIECE, ' + Pf + 'E_NUMLIGNE';

  }

  result := St ;

end;     
*)


procedure TOF_CPVALIDEECR.FormKeyDown( Sender : TObject; var Key: Word; Shift: TShiftState );
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

Initialization
  registerclasses ( [ TOF_CPVALIDEECR ] ) ;
end.


