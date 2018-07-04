{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/12/2001
Modifié le ... : 20/02/2007
Description .. : Source TOF de la FICHE : TOFMULEMPRUNT ()
Mots clefs ... : TOF;TOFMULEMPRUNT
*****************************************************************}
Unit TOFMULEMPRUNT;

Interface

Uses utobdebug,
     StdCtrls,
     Controls,
     Classes,
//     Spin,

{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
{$IFDEF VER150}
     Variants,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     ExtCtrls,
     HEnt1,
     HTB97,
     HMsgBox,
    {$ifdef eAGLClient}
     MaineAGL,emul,
     {$ELSE}
     FE_Main,mul,
     {$ENDIF}
     HDB,
     UTOF,
     TOFEMPIMP,
     TOMEMPRUNT,
     UObjetEmprunt,
     Windows,         // VK_
     Menus,           // TPopUpMenu;
     UlibPieceCompta,
     uLibECriture,
     UTob,
     AGLInit,               // TheMulQ
     HStatus,
     DelVisuE,              // VisuPiecesGenere
     Ed_Tools,              // Pour le videListe
     FpEcrCreParam_Tof,     // Paramètres de génération
     FpRapproCRE_Tof,       // Rapprochement Compta/CRE
     ParamSoc,
     Math,                  // Min
     Ent1,
     SaisUtil;              // Ecrgen


Type
  TOF_MULEMPRUNT = Class (TOF)
    private
       fBoModeSelection : Boolean;
       fDateDebGen : TDateTime;
       POPF11   : TPopUpMenu;
       POPUPTR  : TPopUpMenu;
       POPUPEDIT: TPopUpMenu;
       FEcran   : TFMul ;
      {$IFDEF EAGLCLIENT}
       FListe : THGrid ;
      {$ELSE}
       FListe : THDBGrid ;
      {$ENDIF}
       CalcSimul : TCheckBox;
       XX_WHERE  : THEdit ;

       procedure bRapproClick(Sender : TObject);
       procedure bImprimerTEMClick(Sender : TObject);
       procedure bImprimerTFFClick(Sender : TObject);
       procedure bImprimerPREClick(Sender : TObject);
       procedure bImprimerAMOClick(Sender : TObject);
       procedure bImprimerECRClick(Sender : TObject);
       procedure OnInsertClick(Sender : TObject);
       procedure bDeleteClick(Sender : TObject);
       procedure CalcSimulClick(Sender : TObject);
       procedure bFraisFinaClick(Sender : TObject);
       procedure bRegulFinExeClick(Sender : TObject);
       procedure RechercherDates(Code : integer ; TypeGene : String ; date_fin : TDateTime ; var Date_eche, Date_decal : TDateTime);
//       function  GetField(pStFieldName : String) : Variant;
       procedure grListDblClick(Sender : TObject);
       procedure Ouvrir;
       function  CreerPieceC ( vInfoEcr : TInfoEcriture ; JalGene : String ) : TPieceCompta ;
       function  GenererPieceCompta (TobEmp : Tob ; PieCpt   : TPieceCompta ; TypeGene, JalGene : String ; Mt_inter, Mt_assur : real ; DateFin : TDateTime) : Tob ;
       procedure Genere_regul(JalGene :String ; TobEmp : Tob ; CalcInter, CalcAssur : boolean ; Date_fin : TDateTime ; ListePieces : TList);
       procedure LigneTPieceCreer ( vPieceCompta : TPieceCompta ; JalGene, Sens, Compte_Gene, Libelle, Refce : string ; Montant : real ) ;
       procedure LigneTPieceSolde ( vPieceCompta : TPieceCompta ; JalGene, Compte_Gene, Libelle, Refce : string  ) ;
       function RemplirTob : Tob;
       procedure AddFille (T : TOB ; Q : TQuery);
       Function VerifDates (CodeEmprunt : Integer ; var DateDeb : TDateTime ; var DateFin : TDateTime) : Boolean;
       function MajDateGene(CodeEmprunt : Integer ; DateFin : TDateTime ) : boolean;

    published
       procedure OnNew                    ; override ;
       procedure OnDelete                 ; override ;
       procedure OnClickBDelete(Sender: TObject);
       procedure SuppressionCompte;
       procedure OnUpdate                 ; override ;
       procedure OnLoad                   ; override ;
       procedure OnArgument (S : String ) ; override ;
       procedure OnClose                  ; override ;
       procedure OnPopUpPopF11    (Sender : TObject ); //override;
       procedure OnKeyDownEcran(Sender : TObject ; var Key : Word ; Shift : TShiftState);
  end ;


function CPLanceFiche_Emprunt ( vStParam: string = '') : string ;

Implementation

Function CPLanceFiche_Emprunt ( vStParam: string = '') : string ;
begin
   Result := AGLLanceFiche('FP','FMULEMPRUNT','','', vStParam);
end;

// ----------------------------------------------------------------------------
// Nom    : OnArgument
// Date   : 05/11/2002
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOF_MULEMPRUNT.OnArgument (S : String ) ;
var
  lStNumCompte, lstTypeCompte : String;
begin
  Inherited ;

 {$IFDEF EAGLCLIENT}
  FListe  := THGrid( GetControl('FListe',True) ) ;
 {$ELSE}
  FListe  := THDBGrid( GetControl('FListe',True)) ;
 {$ENDIF}

  // Filtre sur le numéro de compte
  lStNumCompte := ReadTokenSt(S); 
  lStTypeCompte := ReadTokenSt(S);
  {22.05.07 YMO FQ20250}
  // Mode sélection ou ouverture
  fBoModeSelection := S <> '';

  FEcran := TFMul(Ecran) ;

  // Evènements
  FListe.OnDblClick                                  := grListDblClick;
  TToolBarButton97(GetControl('BInsert')).OnClick    := OnInsertClick;
  TToolBarButton97(GetControl('bDelete')).OnClick    := bDeleteClick;
  // Sélection uniquement ?
  SetControlVisible('bInsert',   Not fBoModeSelection);
  SetControlVisible('bdelete',   Not fBoModeSelection);
  SetControlVisible('bImprimer', Not fBoModeSelection);

  CalcSimul:= TCheckBox(GetControl('CALCSIMUL'));
  CalcSimul.State:= cbGrayed;
  CalcSimul.OnClick:=CalcSimulClick;
  XX_WHERE := THEdit(GetControl('XX_WHERE', True)) ;
  XX_WHERE.Text :='' ;
  // Init des controles
 // SetControlText('EMP_CODEEMPRUNT','' );
  SetControlProperty('EMP_CODEEMPRUNT','Value',-1);

  Ecran.OnKeyDown := OnKeyDownEcran;

  POPF11 := TPopUpMenu(GetControl('POPF11',True));
  POPUPTR := TPopUpMenu(GetControl('POPUPTRAITEMENT',True));
  POPUPEDIT := TPopUpMenu(GetControl('POPUPEDITIONS',True));

  POPUPTR.Items[0].OnClick := bFraisFinaClick;
  POPUPTR.Items[1].OnClick := bRegulFinExeClick;

  POPF11.Items[0].OnClick := OnInsertClick;
  POPF11.Items[1].OnClick := bDeleteClick;
  POPF11.Items[2].OnClick := grListDblClick;
  POPF11.Items[3].OnClick := bRapproClick;
  POPF11.Items[4].OnClick := bFraisFinaClick;
  POPF11.Items[5].OnClick := bRegulFinExeClick;
  POPF11.Items[6].OnClick := bImprimerTEMClick;  //TEM
  POPF11.Items[7].OnClick := bImprimerTFFClick;  //TFF
  POPF11.Items[8].OnClick := bImprimerPREClick;  //PRE
  POPF11.Items[9].OnClick := bImprimerAMOClick;  //AMO
  POPF11.Items[10].OnClick := bImprimerECRClick; //ECR

  {FQ19983 18/04/07 YMO Bouton impressions}
  POPUPEDIT.Items[0].OnClick := bImprimerTEMClick;  //TEM
  POPUPEDIT.Items[1].OnClick := bImprimerTFFClick;  //TFF
  POPUPEDIT.Items[2].OnClick := bImprimerPREClick;  //PRE
  POPUPEDIT.Items[3].OnClick := bImprimerAMOClick;  //AMO
  POPUPEDIT.Items[4].OnClick := bImprimerECRClick; //ECR

{  POPF11.Items[3].OnClick := OnClickBInfo;

    // a virer quand fiche dans SOCREF
  if PopF11.Items.Count = 5 then
      POPF11.Items[4].OnClick := OnClickBEtatRappro;
 }
  AddMenuPop(PopF11, '', '');
//  POPF11.OnPopup          := OnPopUpPopF11;

  //FQ 19984 17/04/2007 YMO
  If (Copy(lStNumCompte, 1, 4) = '6611') then lStTypeCompte:='EMP_CPTFRAISFINA';

  if lStNumCompte <> '' then
  begin
     if (lStTypeCompte='') then
     begin
        {renseignement des comptes à partir des paramsoc si on n'a pas déjà un type de compte en paramètre}
         If lStNumCompte=GetParamSocSecur ('SO_CPINTERCOUR', False) then lStTypeCompte:='EMP_CPTINTERCOUR'
        else
         If lStNumCompte=GetParamSocSecur ('SO_CPCHARGAVAN', False) then lStTypeCompte:='EMP_CPTCHARGAVAN'
        else
         If lStNumCompte=GetParamSocSecur ('SO_CPFRAISFINA', False) then lStTypeCompte:='EMP_CPTFRAISFINA'
        else
         If lStNumCompte=GetParamSocSecur ('SO_CPASSURANCE', False) then lStTypeCompte:='EMP_CPTASSURANCE'
        else
         lStTypeCompte:='EMP_NUMCOMPTE';
     end;

     If lStTypeCompte<>'EMP_NUMCOMPTE' then
        TPageControl(GetControl('PAGES', true)).ActivePage := TTabSheet(GetControl('PCOMPTES_ASSOCIES',True));

     SetControlText(lStTypeCompte, lStNumCompte);

  end;

end ;



// ----------------------------------------------------------------------------
// Nom    : OnNew
// Date   : 05/11/2002
// Auteur : D. ZEDIAR
// Objet  : Nouvel emprunt
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOF_MULEMPRUNT.OnNew ;
begin
  Inherited ;
end ;

// ----------------------------------------------------------------------------
// Nom    : OnDelete
// Date   : 05/11/2002
// Auteur : D. ZEDIAR
// Objet  : suppression d'un emprunt
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOF_MULEMPRUNT.OnDelete ;
begin
   Inherited ;
end ;

procedure TOF_MULEMPRUNT.bDeleteClick(Sender : TObject);
begin
  inherited;
  if not BlocageMonoPoste(True) then Exit;

  try

    if (FListe.AllSelected) or (FListe.nbSelected <> 0) then
    begin
      if PgiAsk('Confirmez vous la suppression des emprunts sélectionnés ?', Ecran.Caption) = MrYes then
      begin
        if Transactions(SuppressionCompte, 3) <> oeOk then
          MessageAlerte('Traitement annulé. Erreur lors de la suppression des emprunts.');
      end;
    end
    else
      PgiInfo('Vous n''avez pas sélectionné d''emprunt.', Ecran.Caption);
    FListe.ClearSelected;

  finally
    DeblocageMonoPoste(True);
    TFMul(Ecran).BChercheClick(nil);
  end;
end;

procedure TOF_MULEMPRUNT.SuppressionCompte;
var
  i: integer;
  LeCompte: string;
  fOmEmprunt : TObjetEmprunt;
begin                        

  if FListe.AllSelected then
  begin // On travaille sur tous les enregistrements de la Query
    TheMulQ.First;

    while not TheMulQ.Eof do
    begin
      LeCompte := TheMulQ.FindField('EMP_CODEEMPRUNT').AsString;
      fOmEmprunt := TObjetEmprunt.CreerEmprunt;
     try
      fOmEmprunt.ValeurPK := GetField('EMP_CODEEMPRUNT');
      if transactions(fOmEmprunt.DBSupprimerEmprunt,1) <> oeOk then
         PGIError('Une erreur est survenue à la suppression de l''emprunt !', TitreHalley);
     finally
      fOmEmprunt.Free;
     end;

      TheMulQ.Next;
    end;
  end
  else
    for i := 0 to FListe.NbSelected - 1 do
    begin
      FListe.GotoLeBookMark(i);
      LeCompte := GetField('EMP_CODEEMPRUNT');
         fOmEmprunt := TObjetEmprunt.CreerEmprunt;
     try
      fOmEmprunt.ValeurPK := GetField('EMP_CODEEMPRUNT');

      if transactions(fOmEmprunt.DBSupprimerEmprunt,1) <> oeOk then
         PGIError('Une erreur est survenue à la suppression de l''emprunt '+LeCompte, TitreHalley);
     finally
      fOmEmprunt.Free;
     end;
    end;

end;

// ----------------------------------------------------------------------------
// Nom    : OnUpdate
// Date   : 05/11/2002
// Auteur : D. ZEDIAR
// Objet  : Sélection d'un emprunt
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOF_MULEMPRUNT.OnUpdate ;
begin
   Inherited ;
end ;

// ----------------------------------------------------------------------------
// Nom    : OnLoad
// Date   : 05/11/2002
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOF_MULEMPRUNT.OnLoad ;
begin
  inherited;
end;

procedure TOF_MULEMPRUNT.grListDblClick(Sender : TObject);
begin
   Ouvrir;
end;

procedure TOF_MULEMPRUNT.Ouvrir;
var
  lVaCodeEmprunt : Variant;
begin
  lVaCodeEmprunt := GetField('EMP_CODEEMPRUNT');
  if fBoModeSelection then
  begin
     // En mode sélection, renvoie le code et le libellé de l'emprunt
     if VarIsEmpty(lVaCodeEmprunt) or VarIsNull(lVaCodeEmprunt) then
        TFMul(Ecran).Retour := ''
     else
        TFMul(Ecran).Retour := IntToStr(lVaCodeEmprunt) + ';' + GetField('EMP_LIBEMPRUNT');
     Ecran.Close;
  end
  else
  begin
     // En mode ouverture, lance la fenêtre de saisie des emprunts
     if not (VarIsEmpty(lVaCodeEmprunt) or VarIsNull(lVaCodeEmprunt)) then
     begin
        SaisieEmprunt(lVaCodeEmprunt);
        TFMUL(ECRAN).BChercheClick(nil);
     end;
  end;
end;

procedure TOF_MULEMPRUNT.bRapproClick(Sender : TObject);
begin
  EtatRapproCRECompta;
end;

procedure TOF_MULEMPRUNT.bImprimerTEMClick(Sender : TObject);
begin
  TheMulQ:=nil;
  if not VarIsEmpty(GetField('EMP_CODEEMPRUNT')) then
     ImprimeEmprunt(0,'',False,'TEM');
  {FQ20223 23.05.07 YMO affichage de tous les emprunts
  ImprimeEmprunt(GetField('EMP_CODEEMPRUNT'),GetField('EMP_LIBEMPRUNT'),False,'TEM');}
end;

procedure TOF_MULEMPRUNT.bImprimerTFFClick(Sender : TObject);
begin
  if not (VarIsEmpty(GetField('EMP_CODEEMPRUNT')) Or VarIsEmpty(GetField('EMP_LIBEMPRUNT')))  then
     ImprimeEmprunt(0,'',False,'TFF');
  {FQ20223 23.05.07 YMO affichage de tous les emprunts
  ImprimeEmprunt(GetField('EMP_CODEEMPRUNT'),GetField('EMP_LIBEMPRUNT'),False,'TFF');}
end;

procedure TOF_MULEMPRUNT.bImprimerPREClick(Sender : TObject);
begin
  if not VarIsEmpty(GetField('EMP_CODEEMPRUNT')) then
     ImprimeEmprunt(0,'',False,'PRE');
  {FQ20223 23.05.07 YMO affichage de tous les emprunts
  ImprimeEmprunt(GetField('EMP_CODEEMPRUNT'),GetField('EMP_LIBEMPRUNT'),False,'PRE');}
end;

procedure TOF_MULEMPRUNT.bImprimerAMOClick(Sender : TObject);
begin
  if not VarIsEmpty(GetField('EMP_CODEEMPRUNT')) then
  {FQ20223 23.05.07 YMO affichage de tous les emprunts}
  {31.05.07 YMO Pas nécessaire pour AMO : un seul emprunt}
  ImprimeEmprunt(GetField('EMP_CODEEMPRUNT'),GetField('EMP_LIBEMPRUNT'),False,'AMO');
end;

procedure TOF_MULEMPRUNT.bImprimerECRClick(Sender : TObject);
begin
  if not VarIsEmpty(GetField('EMP_CODEEMPRUNT')) then
     ImprimeEmprunt(0,'',False,'ECR');
  {FQ20223 23.05.07 YMO affichage de tous les emprunts
  ImprimeEmprunt(GetField('EMP_CODEEMPRUNT'),GetField('EMP_LIBEMPRUNT'),False,'ECR');}
end;

procedure TOF_MULEMPRUNT.OnInsertClick(Sender : TObject);
begin
  inherited;
  if GetControlEnabled('EMP_NUMCOMPTE') then
     NouvelEmprunt
  else
     NouvelEmprunt(GetControlText('EMP_NUMCOMPTE'));

  TFMUL(ECRAN).BChercheClick(Nil);
end;


procedure TOF_MULEMPRUNT.OnClickBDelete(Sender: TObject);
var
   fOmEmprunt : TObjetEmprunt;
begin



   if PGIAsk('Etes-vous certain de vouloir supprimer cet emprunt ?', TitreHalley)  = mrNo then
      Exit;

   if VarIsEmpty(GetField('EMP_CODEEMPRUNT')) then
      Exit;

   //
//   if PGIAsk('Confirmez-vous la suppression de cet emprunt ?',TitreHalley) = mrNo then
//      Exit;

   fOmEmprunt := TObjetEmprunt.CreerEmprunt;

   try
      fOmEmprunt.ValeurPK := GetField('EMP_CODEEMPRUNT');

      //
      if transactions(fOmEmprunt.DBSupprimerEmprunt,1) = oeOk then
         TFMUL(ECRAN).BChercheClick(nil)
      else
         PGIError('Une erreur est survenue à la suppression de l''emprunt !', TitreHalley);

   finally
      fOmEmprunt.Free;
   end;
end ;

procedure TOF_MULEMPRUNT.CalcSimulClick(Sender : TObject);
begin
  If CalcSimul.Checked then
    XX_WHERE.Text:=' AND EMP_STATUT=4'
  else
  If (not CalcSimul.Checked) then
    XX_WHERE.Text:=' AND EMP_STATUT<>4'
  else
    XX_WHERE.Text:='';
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 02/03/2007
Modifié le ... :   /  /
Description .. : Point d'entrée pour le calcul des frais financiers
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TOF_MULEMPRUNT.bFraisFinaClick(Sender : TObject);
var
  TobEmp, TInfo, TPieGene, TPieTemp : Tob ;
  CalcInter, CalcAssur : Boolean ;
  T : TobPiececompta ;
  Q, QP : TQuery ;
  i, NumEmprunt : integer ;
  DateDeb, DateFin : TDateTime ;
  Mt_inter, Mt_assur : Real ;
  JalGene, TypeGene : String ; // A Avance ; R Retard ; F Frais financiers
  ListePieces : TList ;
  VerifOK : Boolean ;
  PieCpt : TPieceCompta ;
  InfoEcr  : TInfoEcriture ;
  modepiece : boolean ;
begin

  TobEmp:=RemplirTob;
  If TobEmp.Detail.Count=0 Then
  begin
   PGIInfo(TraduireMemoire('Veuillez sélectionner au moins une ligne d''emprunt'),'');
   LastError := -1;
   Exit;
  end;

  TypeGene:='F' ;
  Mt_Inter:=0;
  Mt_Assur:=0;
  TInfo := TOB.Create('', nil, -1);
  TheTOB := TInfo;
  CPLanceFiche_ParamCRE(TypeGene+';'+DateToStr(fDateDebGen));
  TInfo.Free;
  if TheTOB = nil then
  begin
    TFMul(Ecran).BChercheClick(nil);
    exit;
  end;

  DateDeb := StrToDate(TheTOB.GetString('EMP_DATEDEB'));
  DateFin  := StrToDate(TheTOB.GetString('EMP_DATEFIN'));
  CalcInter := StrToBool(TheTOB.GetString('EMP_INTER'));
  CalcAssur := StrToBool(TheTOB.GetString('EMP_ASSUR'));
  JalGene := TheTOB.GetString('EMP_JOURNAL');

  ListePieces := TList.Create ;

  VerifOk:=True;
  for i:=0 to TobEmp.Detail.Count-1 do
  begin
    If not VerifDates(TobEmp.Detail[i].GetInteger('EMP_CODEEMPRUNT'), DateDeb, DateFin) then
    begin
      VerifOk:=False;
      Break;
    end;
  end;

  If (not VerifOk)
  and (PGIAsk ('"Attention, vous allez générer des écritures sur une période déjà concernée par une précédente régularisation. Voulez vous continuer ?', 'Dates')=mrNo) then
    exit;

    ////////////////////
    T := TobPieceCompta.Create('', Nil, -1);
    {Création}
    InfoEcr := T.CreateInfoEcr('') ;
    /////////////////////

    modepiece:=false;
    Q:=OpenSQL('SELECT J_MODESAISIE FROM JOURNAL WHERE J_JOURNAL="'+UpperCase(JalGene)+'"', TRUE) ;
    if not Q.EOF then modepiece:=(Q.FindField('J_MODESAISIE').AsString='-');

    PieCpt:=nil;
    TPieGene:=nil;
for i:=0 to TobEmp.Detail.Count-1 do
begin

    {FQ20663 13.06.07 YMO On crée une pièce pour chaque emprunt, et en mode bordereau un seul folio}
    If (i=0) or (modepiece) then PieCpt := CreerPieceC(InfoEcr, JalGene) ;

    NumEmprunt:=TobEmp.Detail[i].GetInteger('EMP_CODEEMPRUNT');

    {On réduit à la période non encore générée}
    {28.05.07 YMO On laisse libre l'utilisateur de générer plusieurs fois sur la même période
    VerifDates(NumEmprunt, DateDeb, DateFin) ;
    }

    {Calcul Montant intérêts & assurance}
    QP := OpenSQL ('SELECT SUM(ECH_INTERET) SOMMINT, SUM(ECH_ASSURANCE) SOMMASSU'
                  +' FROM FECHEANCE WHERE ECH_CODEEMPRUNT="'+intToStr(NumEmprunt)
                  +'" AND ECH_DATE>="'+UsDateTime(DateDeb)
                  +'" AND ECH_DATE<="'+UsDateTime(DateFin)
                  +'"', TRUE);
    if not QP.EOF then
    begin
      If CalcInter then Mt_inter := QP.FindField('SOMMINT').AsFloat ;
      If CalcAssur then Mt_assur := QP.FindField('SOMMASSU').AsFloat ;
    end ;
    Ferme (QP);

    If (mt_inter<>0) Or (mt_assur<>0) then
       TPieTemp := GenererPieceCompta(TobEmp.Detail[i], PieCpt, TypeGene, JalGene, Mt_inter, Mt_assur, Datefin)
    else
       TPieTemp := nil;

    {15.06.07 Dans le cas des bordereaux, on ne réinitialise pas l'objet unique TPieGene}
    If Modepiece or (TPieTemp<>nil) then TPieGene:=TPieTemp;

     {numéro de pièce pour les bordereaux}
    if PieCpt.ModeSaisie<>msPiece then PieCpt.AttribNumeroDef ;

    {FQ20663 13.06.07 YMO On enregistre une pièce pour chaque emprunt, et en mode bordereau uniquement sur le dernier emprunt (même folio)}
    if (PieCpt.ModeSaisie=msPiece) or (i=TobEmp.Detail.Count-1) then
    begin

        {validation} {15.06.07  YMO Pas de message d'erreur autres que ceux de ULibPieceCompta}
        if (PieCpt <> nil) or (PieCpt.Detail.Count <> 0) and (PieCpt.isValidPiece) then
           PieCpt.Save;
{        else
           PgiBox('Erreur à la création de la pièce'); }

        If (ListePieces<>nil) And (TPieGene<>nil) then
        begin
          {Mise à jour de la date de dernière génération d'écritures}
          if not MajDateGene(NumEmprunt, DateFin) then
          begin
            PgiInfo('Traitement annulé. Erreur lors de la mise à jour de l''emprunt.', '');
            LastError := -1;
            Exit;
          end;

          {FQ20542  07.06.07  YMO Pb message d'erreur : indice hors limite}
          ListePieces.Add(TPieGene.Detail[0]) ;
        end;
    end;
end;

  If ListePieces.Count>0 then
      VisuPiecesGenere(ListePieces,EcrGen, 19)
  else
    PgiInfo('Aucune écriture générée');

  VideListe(ListePieces) ;
  ListePieces.Free ;

  TFMul(Ecran).BChercheClick(nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 02/03/2007
Modifié le ... :   /  /
Description .. : On recherche les dates d'échéance qui servent de base
Suite ........ : pour le calcul des régularisations
Mots clefs ... :
*****************************************************************}
procedure TOF_MULEMPRUNT.RechercherDates(Code : integer ; TypeGene : String ; date_fin : TDateTime ; var Date_eche, Date_decal : TDateTime);
var
   lStSQL : String ;
   lTbEcheances, lTbDetail : Tob ;
   lEche : Integer ;
begin

    // Extraction des données de l'emprunt
    lStSQL := 'SELECT * FROM FEMPRUNT, FECHEANCE';
    lStSQL := lStSQL + ' WHERE ECH_CODEEMPRUNT = EMP_CODEEMPRUNT';
    lStSQL := lStSQL + ' AND EMP_CODEEMPRUNT ="'+intTostr(Code)+ '"';
    lStSQL := lStSQL + ' ORDER BY ECH_CODEEMPRUNT, ECH_DATE';

    lTbEcheances := Tob.Create('',Nil,-1);

    lTbEcheances.LoadDetailFromSQL(lStSQL);

      // Boucle sur les échéances
      For lEche:= 0 to lTbEcheances.Detail.Count- 1 do
      begin

         lTbDetail := lTbEcheances.Detail[lEche];
         If TypeGene='A' then {CCA Charges comptabilisées d'avance}
         begin
           If lTbDetail.GetDateTime('ECH_DATE')<=Date_fin then
              date_eche:= lTbDetail.GetDateTime('ECH_DATE');
           If lTbDetail.GetDateTime('ECH_DATE')> Date_fin then
           begin
              date_decal:= lTbDetail.GetDateTime('ECH_DATE');
              break;
           end;
         end
         else
         begin                {Intérêts courus}
           If lTbDetail.GetDateTime('ECH_DATE')< Date_fin then
              date_decal:= lTbDetail.GetDateTime('ECH_DATE');
           If lTbDetail.GetDateTime('ECH_DATE')>=Date_fin then
           begin
              date_eche:= lTbDetail.GetDateTime('ECH_DATE');
              break;
           end;
         end;

      end;

      lTbEcheances.Free;

end;



{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 02/03/2007
Modifié le ... :   /  /
Description .. : Régularisation en fin de période : exercice comptable ou
Suite ........ : situation intermédiaire
Mots clefs ... :
*****************************************************************}
procedure TOF_MULEMPRUNT.bRegulFinExeClick(Sender : TObject);
var
  TobEmp, TInfo : Tob ;
  Date_fin : TDateTime ;
  JalGene, TypeGene : String ;       // A Avance ; R Retard ; F Frais financiers
  CalcInter, CalcAssur : Boolean ;
  ListePieces : TList ;
begin

  TobEmp:=RemplirTob;
  If TobEmp.Detail.Count=0 Then
  begin
   PGIInfo(TraduireMemoire('Veuillez sélectionner au moins une ligne d''emprunt'),'');
   LastError := -1;
   Exit;
  end;

  TInfo := TOB.Create('', nil, -1);

  TheTOB := TInfo;

  CPLanceFiche_ParamCRE(TypeGene);
  TInfo.Free;
  if TheTOB = nil then
  begin
    TFMul(Ecran).BChercheClick(nil);
    exit;
  end;

  Date_fin  := StrToDate(TheTOB.GetString('EMP_DATEFIN'));
  CalcInter := StrToBool(TheTOB.GetString('EMP_INTER'));
  CalcAssur := StrToBool(TheTOB.GetString('EMP_ASSUR'));
  JalGene := TheTOB.GetString('EMP_JOURNAL');

  ListePieces := TList.Create ;

  Genere_regul(JalGene, TobEmp, CalcInter, CalcAssur, Date_fin, ListePieces);

  If ListePieces.Count> 0 then
      VisuPiecesGenere(ListePieces,EcrGen, 19)
  else
    PgiInfo('Aucune écriture générée');

  VideListe(ListePieces) ;
  ListePieces.Free ;

  TFMul(Ecran).BChercheClick(nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 16/03/2007
Modifié le ... :   /  /
Description .. :
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TOF_MULEMPRUNT.Genere_regul(JalGene : String ; TobEmp:  Tob ; CalcInter, CalcAssur : boolean ; Date_fin : TDateTime ; ListePieces : TList);
var
  T : TobPiececompta ;
  TPieGene, TPieTemp : Tob ;
  Periode, Quota : Double ;
  i, numemprunt : Integer ;
  Date_eche, Date_decal : TDateTime ;
  Q, QP : TQuery ;
  Mt_inter, Mt_assur : Real ;
  TypeGene : String;
  PieCpt : TPieceCompta ;
  InfoEcr  : TInfoEcriture ;
  modepiece : boolean ;
begin

////////////////////
T := TobPieceCompta.Create('', Nil, -1);
{Création}
InfoEcr := T.CreateInfoEcr('') ;
/////////////////////
modepiece:=false;
Q:=OpenSQL('SELECT J_MODESAISIE FROM JOURNAL WHERE J_JOURNAL="'+UpperCase(JalGene)+'"', TRUE) ;
if not Q.EOF then modepiece:=(Q.FindField('J_MODESAISIE').AsString='-');

PieCpt:=nil;
TPieGene:=nil;

for i:=0 to TobEmp.Detail.Count-1 do
begin

    NumEmprunt:=TobEmp.Detail[i].GetInteger('EMP_CODEEMPRUNT');

    {FQ20663 13.06.07 YMO On crée une pièce pour chaque emprunt, et en mode bordereau un seul folio}
    If (i=0) or (modepiece) then PieCpt := CreerPieceC(InfoEcr, JalGene) ;

    {Détermination du type de régularisation}
    QP := OpenSQL ('SELECT EMP_DATECONTRAT, EMP_DATEDEBUT FROM FEMPRUNT '
                  +'WHERE EMP_CODEEMPRUNT="'+IntToStr(NumEmprunt)+'"'
                  +'AND EMP_DATEDEBUT<="'+USDateTime(Date_fin)+'"', TRUE); {FQ20543 07.06.07  YMO}

    TypeGene:='';
    if not QP.EOF then
    begin
      If QP.FindField('EMP_DATECONTRAT').AsString < QP.FindField('EMP_DATEDEBUT').AsString Then
        TypeGene:='R'
      else
        TypeGene:='A';
    end ;
    Ferme (QP);

    RechercherDates(NumEmprunt, TypeGene, Date_fin, Date_eche, Date_decal);

    Periode:=Abs(Date_decal-Date_eche);
    If Periode =0 then Periode :=1;

    If TypeGene='A' then {CCA Avance}
      // On enlève la partie qu'on a déjà payée, de la date choisie jusqu'à la prochaine échéance
      Quota:=Date_decal-Date_fin
    else                 {IC Retard}
      // On enlève la partie qu'on n'a pas encore payée, de la précédente échéance à la date choisie
      Quota:=Date_fin-Date_decal;

    {Calcul Montant intérêts & assurance}
    QP := OpenSQL ('SELECT SUM(ECH_INTERET) SOMMINT, SUM(ECH_ASSURANCE) SOMMASSU'
                  +' FROM FECHEANCE WHERE ECH_CODEEMPRUNT="'+intToStr(NumEmprunt)
                  +'" AND ECH_DATE="'+UsDateTime(Date_eche)+'"', TRUE);

    Mt_inter:=0; Mt_assur:=0;
    if not QP.EOF then
    begin
      If CalcInter then Mt_inter := QP.FindField('SOMMINT').AsFloat ;
      If CalcAssur then Mt_assur := QP.FindField('SOMMASSU').AsFloat ;
    end ;
    Ferme (QP);

    {Calcul du prorata}
    Mt_inter:=Mt_inter*Quota/Periode ;
    Mt_assur:=Mt_assur*Quota/Periode ;
    //JMax:=StrToInt(FormatDateTime('d',FinDeMois(EncodeDate(an,mois,1)))) ;

    If (TypeGene<>'') {FQ20543 15.06.07  YMO Complément : pas de génération si emprunt popstérieur à la date de génération}
    and ((mt_inter<>0) Or (mt_assur<>0)) then
       TPieTemp := GenererPieceCompta(TobEmp.Detail[i], PieCpt, TypeGene, JalGene, Mt_inter, Mt_assur, Date_Fin)
    else
       TPieTemp := nil;

    {15.06.07 Dans le cas des bordereaux, on ne réinitialise pas l'objet unique TPieGene}
    If Modepiece or (TPieTemp<>nil) then TPieGene:=TPieTemp;

    {FQ20663 13.06.07 YMO On enregistre une pièce pour chaque emprunt, et en mode bordereau uniquement sur le dernier emprunt (même folio)}
    if (PieCpt.ModeSaisie=msPiece) or (i=TobEmp.Detail.Count-1) then
    begin

        {validation} {15.06.07  YMO Pas de message d'erreur autres que ceux de ULibPieceCompta}
        if (PieCpt <> nil) or (PieCpt.Detail.Count <> 0) and (PieCpt.isValidPiece) then
           PieCpt.Save;

        If (ListePieces<>nil) And (TPieGene<>nil) then
        begin
          {Mise à jour de la date de dernière génération d'écritures}
          if not MajDateGene(NumEmprunt, Date_Fin) then
          begin
            PgiInfo('Traitement annulé. Erreur lors de la mise à jour de l''emprunt.', '');
            LastError := -1;
            Exit;
          end;
          {FQ20542  07.06.07  YMO Pb message d'erreur : indice hors limite}
          ListePieces.Add(TPieGene.Detail[0]) ;
        end;
    end;


end;

end;


// ----------------------------------------------------------------------------
// Nom    : OnClose
// Date   : 05/11/2002
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOF_MULEMPRUNT.OnClose ;
begin
//  gObMulEmprunt := Nil;
//  gQrMulEmprunt := Nil;
  Inherited ;
end ;


procedure TOF_MULEMPRUNT.OnKeyDownEcran(Sender : TObject ; var Key : Word ; Shift : TShiftState);

begin
    case key of

        VK_F11 : POPF11.Popup(Mouse.CursorPos.x, Mouse.CursorPos.y);

        VK_DELETE : bDeleteClick (Sender);  {17/04/07 FQ19985 YMO}
    end;

    if ( Ecran <> nil ) and ( Ecran is  TFMul ) then
        TFMul(Ecran).FormKeyDown(Sender,Key,Shift);
end;

procedure TOF_MULEMPRUNT.OnPopUpPopF11(Sender: TObject);
begin
{  PurgePopup( PopF11 );
  AddMenuPop( PopF11, '', '');
  InitPopUp( True );
  inherited; }
end;


// ----------------------------------------------------------------------------
// Nom    :
// Date   : 05/11/2002
// Auteur : D. ZEDIAR
// Objet  : Renvoie la valeur d'un champ de l'enregistrement sélectionné dans
//          la grille
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
{Function TOF_MULEMPRUNT.GetField(pStFieldName : String) : Variant;
begin
   if TQuery(GetControl('Q')).Eof then
      Result := Unassigned
   else
      Result := TQuery(GetControl('Q')).FindField(pStFieldName).Value;
end;
}
{Crée une ligne dans le TPieceCompta}
procedure  TOF_MULEMPRUNT.LigneTPieceCreer ( vPieceCompta : TPieceCompta ; JalGene, Sens, Compte_Gene, Libelle, Refce : string ; Montant : real );
begin

  If Montant<=0 then exit;

  vPieceCompta.newrecord ;

  vPieceCompta.PutValue(vPieceCompta.Detail.Count,'E_GENERAL', Compte_Gene );
  {28.05.07 YMO Mise en place du libellé}
  vPieceCompta.PutValue(vPieceCompta.Detail.Count,'E_LIBELLE', Libelle );

  {FQ21118 23.07.07  YMO}
  vPieceCompta.PutValue(vPieceCompta.Detail.Count,'E_REFINTERNE', Refce );

  If Sens='C' then
    vPieceCompta.PutValue(VpieceCompta.Detail.count,'E_CREDITDEV', FloatToStr(Montant) )
  Else
    vPieceCompta.PutValue(VpieceCompta.Detail.count,'E_DEBITDEV', FloatToStr(Montant) );

end;

{Crée la dernière ligne pour solder le TPieceCompta}
procedure  TOF_MULEMPRUNT.LigneTPieceSolde ( vPieceCompta : TPieceCompta ; JalGene, Compte_Gene, Libelle, Refce : String );
begin

  vPieceCompta.newrecord ;
  vPieceCompta.PutValue(VpieceCompta.Detail.count,'E_GENERAL', Compte_Gene );
  {28.05.07 YMO Mise en place du libellé}
  vPieceCompta.PutValue(VpieceCompta.Detail.count,'E_LIBELLE', Libelle );
  vPieceCompta.AttribSolde(VpieceCompta.Detail.count);
  {FQ21118 23.07.07  YMO}
  vPieceCompta.PutValue(VpieceCompta.Detail.count,'E_REFINTERNE', Refce );
end;


{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 02/03/2007
Modifié le ... :   /  /
Description .. : Initialise la saisie d'un TPieceCompta
Mots clefs ... :
*****************************************************************}
function  TOF_MULEMPRUNT.CreerPieceC ( vInfoEcr : TInfoEcriture ; JalGene : String ) : TPieceCompta ;
begin

  result   := TPieceCompta.CreerPiece( vInfoEcr ) ;

  result.SetMultiEcheOff ; // pas de multi-échéance
  result.Contexte.AttribRIBOff := True ; // pas d'attribution auto du rib :

  // Entête {25.05.07 FQ20399 YMO Passage du journal}
  result.InitPiece( JalGene, V_PGI.DateEntree ) ;
  result.InitSaisie ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 02/03/2007
Modifié le ... :   /  /
Description .. : Génération d'un TPieceCompta
Mots clefs ... :
*****************************************************************}
function  TOF_MULEMPRUNT.GenererPieceCompta (TobEmp : Tob ; PieCpt : TPieceCompta ; TypeGene, JalGene : String ; Mt_inter, Mt_assur : real ; DateFin : TDateTime) : Tob ;
var
    Cpt_inter, Cpt_assur, Cpt_solde, {Compte d'emprunt, compte de contrepartie}
    Sens, Libelle, Refce : String;
    Q : TQuery;
begin

Cpt_solde:=TobEmp.GetString('EMP_NUMCOMPTE') ;
Cpt_inter:=TobEmp.GetString('EMP_CPTFRAISFINA') ;
Cpt_assur:=TobEmp.GetString('EMP_CPTASSURANCE') ;

If TypeGene='F' then {(F)rais financiers sur les échéances}
begin
Libelle:='Frais '+TobEmp.GetString('EMP_LIBEMPRUNT') ;
Cpt_solde:=TobEmp.GetString('EMP_NUMCOMPTE') ;
Sens:='D' ;
end
else
If TypeGene='A' then {Régularisation de fin de période : (A)vance}
begin
Libelle:='Régularisation '+TobEmp.GetString('EMP_LIBEMPRUNT') ;
Cpt_solde:=TobEmp.GetString('EMP_CPTCHARGAVAN') ;
Sens:='C' ;
end
else
If TypeGene='R' then {Régularisation de fin de période : (R)etard}
begin
Libelle:='Régularisation '+TobEmp.GetString('EMP_LIBEMPRUNT') ;
Cpt_solde:=TobEmp.GetString('EMP_CPTINTERCOUR') ;
Sens:='D' ;
end ;

Refce:=TobEmp.GetString('EMP_CODEEMPRUNT') ;

{FQ21225 07/11/2007 YMO Pas de génération sur des comptes visés}
Q:=OpenSQL('SELECT 1 FROM GENERAUX WHERE G_VISAREVISION="X" AND '
          +'((G_GENERAL="'+Cpt_solde+'") OR (G_GENERAL="'
          +Cpt_inter+'") OR (G_GENERAL="'+Cpt_assur+'"))', TRUE) ;

{Pas de message : traitement en masse possible}
If (not Q.EOF) or (Cpt_solde='') or (Cpt_inter='') or (Cpt_assur='') then
begin
  Result:=nil;
  exit;
end;

{FQ20542  15.06.07  YMO Renseignement de la date de fin de période sur la pièce générée}
PieCpt.PutEntete('E_DATECOMPTABLE', DateFin);

{1ere ligne : intérêts}
LigneTPieceCreer (PieCpt, JalGene, Sens, Cpt_inter, Libelle, Refce, Mt_inter);
{1ere ligne : assurance}
LigneTPieceCreer (PieCpt, JalGene, Sens, Cpt_assur, Libelle, Refce, Mt_assur);
{Contrepartie}
LigneTPieceSolde(PieCpt, JalGene, Cpt_solde, Libelle, Refce);

Result:=PieCpt;

end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 02/03/2007
Modifié le ... :   /  /
Description .. : Création de la Tob Emprunt
Mots clefs ... :
*****************************************************************}
function  TOF_MULEMPRUNT.RemplirTob : Tob;
var lMulQ : TQuery ;
    i     : integer;
    TobEmp : Tob ;
begin

 {$IFDEF EAGLCLIENT}
  lMulQ := FEcran.Q.TQ;
  lMulQ.Seek( FListe.Row - 1 );
 {$ELSE}
  lMulQ := FEcran.Q ;
 {$ENDIF}

  TobEmp := Tob.Create('EMPRUNT_', nil, -1) ;

  fDateDebGen:=VH^.EnCours.Deb; {FQ13383 YMO 20/04/07 par défaut dates d'exercice}

  If TFmul(Ecran).FListe.AllSelected then
  begin
      LMulQ.First;

      While not LMulQ.EOF do
      begin
         AddFille(TobEmp, LMulQ);
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

    AddFille(TobEmp, LMulQ);
  end;

  Result:=TobEmp;

  FiniMove;

end;

procedure TOF_MULEMPRUNT.AddFille (T : TOB ; Q : TQuery);
var
      TF : Tob;
      {$IFNDEF EAGLCLIENT}
      i : Integer;
      {$ENDIF}
begin
      TF := Tob.Create('EMPRUNT_', T, -1);
      {$IFDEF EAGLCLIENT}
      // Renseigne les valeurs
      TF.Dupliquer(Q.CurrentFille,true,true);
      {$ELSE}
      // Renseigne les valeurs
      for i := 0 to Q.FieldCount - 1 do
        TF.AddChampSupValeur(Q.Fields.Fields[i].FieldName, Q.Fields.Fields[i].AsString, False);
      {$ENDIF}
      {FQ13383 YMO 20/04/07 par défaut dates d'exercice}
      //Renseignement de la date de génération par défaut
      fDateDebGen:=Max(fDateDebGen, Q.FindField('EMP_DERDATEGENEFF').AsDateTime +1);

end;


{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 02/03/2007
Modifié le ... :   /  /
Description .. : On ne lance pas 2 fois le traitement sur la même période
Suite ........ : On réduit la période à la partie non encore traitée
Mots clefs ... :
*****************************************************************}
Function TOF_MULEMPRUNT.VerifDates (CodeEmprunt : Integer ; var DateDeb : TDateTime ; var DateFin : TDateTime) : Boolean;
var
  QP : TQuery ;
  DateDerGene : TDateTime ;
begin

    Result:=true;
    DateDerGene:=iDate1900;

    {Recherche de la date de dernière modif.}
    QP := OpenSQL ('SELECT EMP_DERDATEGENEFF FROM FEMPRUNT '
                  +'WHERE EMP_CODEEMPRUNT="'+intToStr(CodeEmprunt)
                  +'"', TRUE);
    if (not QP.EOF) and (QP.FindField('EMP_DERDATEGENEFF').AsDateTime<>0) then
         DateDerGene:=QP.FindField('EMP_DERDATEGENEFF').AsDateTime;

    {On ne prend que la partie non encore générée}
    If DateDeb < DateDerGene then
    begin
       //DateDeb:=DateDerGene+1 ;
       Result:=False;
    end;
    {La période est déjà générée : delta de zéro le traitement n'est pas lancé}
    {06.06.07 On laisse le choix
    If DateFin < DateDerGene then DateDeb:=DateFin ;}

    Ferme (QP);

end;

function TOF_MULEMPRUNT.MajDateGene(CodeEmprunt : Integer ; DateFin : TDateTime ) : boolean ;
begin
  Result := True;
  try
    ExecuteSQL('UPDATE FEMPRUNT SET EMP_DERDATEGENEFF ="' +UsDateTime(DateFin)
              +'" WHERE EMP_CODEEMPRUNT="'+intToStr(CodeEmprunt)+'"') ;

  except
    Result := False;
  end;
end;



Initialization
  registerclasses ( [ TOF_MULEMPRUNT ] ) ;
//  gQrMulEmprunt := Nil;
end.





