{ Unité : Source TOF de la FICHE : CPSUPPRUB ()
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
4.3.1.101.001    23/06/03    JP     Création de l'unité
5.X.X            24/02/04    JP     Lors de la suppression d'une rubrique et une seule,
                                    on supprimait toujours la première de la liste
5.0.5.001.001    04/03/04    JP     Mise en place de la Trésorerie
5.0.5.001.001    09/03/04    JP     Gestion du bouton Inserrt pour la Trésorerie
6.xx.xxx.xxx     19/07/04    JP     Gestion des commissions de mouvements pour la tréso
6.20.001.002     02/02/05    VL     Conseils et avertissements + Modif pour les budgets
6.50             25/07/05    MD     Dans l'option "suppression d'une rubrique" (onglet standard)
                                    on affiche des libellés et des champs correspondant à des budgets
                                    FQ 16158
6.50.001.021     03/10/05    JP     FQ TRESO 10264 : Changement du titre de la fiche pour la Tréso
                                    FQ TRESO 10265 : Remise en place du bouton "Ouvrir" et création
                                    de 2 boutons de suppression (1 pour la compta, l'autre pour la Tréso).
6.50.001.025     03/11/05    JP     Modification de l'affectation de la liste pour la Tréso (SetDBListe)
7.00.001.001     27/12/05    JP     FQ TRESO 10308 : Affichage de << Tous >> pour le signe des rubriques
7.03.001.003     08/09/06    CA     FQ 18744 : voir uniquement les rubriques supprimables dans le multicritères
7.06.001.001     09/10/06    JP     FQ TRESO 10363 : Libellé du bouton Ouvrir
7.06.001.001     03/11/06    JP     FQ TRESO 10275 : Contrôle du paramétrage pour la synchronisation
8.10.006.003     08/01/08    JP     FQ 22146 : Mise en place du contrôle de RB_NODOSSIER
--------------------------------------------------------------------------------------}
Unit CPSUPPRUB_TOF ;

Interface

uses StdCtrls,
     Controls,
     Classes,
    {$IFDEF EAGLCLIENT}
     MaineAGL, eMul,
    {$ELSE}
     FE_Main, db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Mul, HDB,
    {$ENDIF}
     forms,
{$IFNDEF TRESO} CpteUtil, {$ENDIF}
     sysutils,
{$IFDEF VER150}
   Variants,
 {$ENDIF}
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Ent1,
     HQry,
     AGLInit,
     UTOF,
     HPanel,
     UTob,
     RUBRIQUE_TOM;

Type
  TOF_CPSUPPRUB = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    FCtxSurRubrique : TCtxSurRubrique;
    Nblig      : Integer ;
    Effacer    : Boolean ;
    NotEffacer : Boolean ;
    GeneCode,
    TypCod     : string ;
    fDateValidite : TDateTime;
    fPredefini : string;
    fNoDossier : string;
    TDelgene : TList;
    TNotDel  : TList;
    LMsg     : THMsgBox;
    {$IFDEF EAGLCLIENT}
    Fliste   : THGrid;
    {$ELSE}
    Fliste   : THDBGrid;
    {$ENDIF}
    Fiche    : TFMul;

    RB_SIGNERUB  : THValComboBox;
    RB_TYPERUB   : THValComboBox;
    TRB_SIGNERUB : THLabel;

    procedure FListeDblClick(Sender: TObject);
    procedure BOuvrirClick  (Sender: TObject);
    procedure BMajRubClick  (Sender: TObject);
  {$IFDEF TRESO}
    procedure BInsertOnClick (Sender : TObject); {JP 09/03/04}
    procedure BControlOnClick(Sender : TObject); {JP 03/11/06 : FQ 10275}
  {$ENDIF}
    procedure ChargeCbTypRub ;
    procedure InitMsg ;
    procedure FermeMsg ;
    procedure Degage ;

    function  Detruit(St,St1, stPredefini, stNoDossier : string; DateValidite : TDateTime) : Byte ;
    function  AfficheMsg (num : Integer; Av, Ap : string ) : Word ;
  end ;

procedure SuppressionRubrique ;
procedure SuppressionRubriqueV2(Budget : Boolean) ;


implementation

uses
  {$IFDEF EAGLCLIENT}
  MenuOLX,
  {$ELSE}
  MenuOLG,
  {$ENDIF EAGLCLIENT}

  {$IFDEF MODENT1}
  CPVersion,
  {$ENDIF MODENT1}

  {$IFDEF TRESO}
  UObjGen, HStatus, TRGrille_TOF, Constantes, 
  {$ELSE}
  RapSuppr,
  {$ENDIF}
  UtilPgi,
  CPVERIFRUBDOSSIER_TOF, {JP 08/01/08 : FQ 22146 : CPLanceFiche_CPVERIFRUBDOSSIER}
  HTB97{FQ TRESO 10265};

{---------------------------------------------------------------------------------------}
procedure SuppressionRubrique ;
{---------------------------------------------------------------------------------------}
begin
  if not _BlocageMonoPoste(False,'',TRUE) then Exit ;
  AglLanceFiche('CP', 'CPSUPPRUB', '', '', 'NON;OUI;');{NON : JP 31/05/05 : FQ 15958}
end ;

{---------------------------------------------------------------------------------------}
procedure SuppressionRubriqueV2(Budget : Boolean) ;
{---------------------------------------------------------------------------------------}
var
  ch : string;
begin
  if not _BlocageMonoPoste(False,'',TRUE) then Exit ;
  if Budget then ch := 'OUI;NON;' //OUI;' // VL 02/02/2005
            else ch := 'NON;NON;'; {OUI : JP 31/05/05 : FQ 15958}
  AglLanceFiche('CP', 'CPSUPPRUB', '', '', ch);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPSUPPRUB.OnLoad ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
//  ChargeCbTypRub ; {JP 04/03/04 : cf le OnArgument}
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPSUPPRUB.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
{$IFNDEF TRESO}
var
  Ok        : Boolean;
  bStdCegid : Boolean;
{$ENDIF TRESO}
begin
  inherited;

  if ReadTokenSt(S) = 'OUI' then
    FCtxSurRubrique := CtxBudget
  else
    FCtxSurRubrique := CtxRubrique;

  Fiche  := TFMul(Ecran);
  {$IFDEF TRESO}
  Fiche.FNomFiltre := 'RUBTRESO';
  {RB_TYPERUB : TR pour Tréso, P (Positif), N (Négatif)}
  SetControlText('XX_WHERE', GetControlText('XX_WHERE') + ' RB_NATRUB = "TRE" AND (RB_TYPERUB = "TRP" OR RB_TYPERUB = "TRN")');
  Ecran.HelpContext := 50000143;
  {$ELSE}
  {JP 31/05/05 : FQ 15958 : True si l'on vient de SuppressionRubrique, à False si l'on vient de SuppressionRubriqueV2}
  Ok     := ReadTokenSt(S) = 'OUI';
  if FCtxSurRubrique <> CtxBudget then
  begin
    Fiche.FNomFiltre := 'SUPPRRUBRI' ;
    SetControlText('XX_WHERE', '(RB_NATRUB="CPT")');
    Ecran.HelpContext := 7790000;
  end
  else
  begin
    Fiche.FNomFiltre := 'SUPPRRUBRIBUD' ;
    SetControlText('XX_WHERE', 'RB_NATRUB="BUD" ');
    Fiche.HelpContext := 15421000 ;
  end ;

  if Ok then begin
    SetControlText('XX_WHERE', '');
    Fiche.FNomFiltre := 'SUPPRRUBRI';
  end
  else begin
    bStdCegid := EstSpecif('51502') and (ctxStandard in V_PGI.PGIContexte);

    { FQ 18744 : voir uniquement les rubriques supprimables dans le multicritères }
    (*
    if (not bStdCegid) then
      SetControlText('XX_WHERE', GetControlText('XX_WHERE') + ' AND RB_RUBRIQUE NOT LIKE "@%"');
    if (not (ctxStandard in V_PGI.PGIContexte)) then
      SetControlText('XX_WHERE', GetControlText('XX_WHERE') + ' AND RB_RUBRIQUE NOT LIKE "$%"');
    *)
    if bStdCegid then
      SetControlText('XX_WHERE', GetControlText('XX_WHERE') + ' AND RB_PREDEFINI="CEG" ')
    else if (ctxStandard in V_PGI.PGIContexte) then
      SetControlText('XX_WHERE', GetControlText('XX_WHERE') + ' AND RB_PREDEFINI="STD" ')
    else SetControlText('XX_WHERE', GetControlText('XX_WHERE') + ' AND RB_PREDEFINI="DOS" ');
  end;
  {$ENDIF}

  TDelgene := TList.Create;
  TNotDel  := TList.Create;

  FListe := Fiche.FListe;
  FListe.OnDblClick := FListeDblClick;
  {JP 03/10/05 : FQ TRESO 10265}
  Fiche.BOuvrir.OnClick := FListeDblClick;
  if Assigned(GetControl('BDELETE')) then
    TToolbarButton97(GetControl('BDELETE')).OnClick := BOuvrirClick;
  if Assigned(GetControl('BDELETE1')) then
    TToolbarButton97(GetControl('BDELETE1')).OnClick := BOuvrirClick;

  InitMsg;

  RB_SIGNERUB  := THValComboBox(GetControl('RB_SIGNERUB'));
  RB_TYPERUB   := THValComboBox(GetControl('RB_TYPERUB'));
  TRB_SIGNERUB := THLabel      (GetControl('TRB_SIGNERUB'));

  {$IFDEF TRESO}
  {Pour la Tréso la gestion des rubriques est simplifiée : elle permet juste de rattacher
   des comptes généraux à des rubriques : les rubriques sont des simili Flux et permettent
   de regrouper les les écritures et de donner un libellé aux groupe de comptes}
  SetControlVisible('RB_TYPERUB'  , False);
  SetControlVisible('TRB_TYPERUB' , False);
  {JP 09/10/06 : FQ TRESO 10363 : On laisse le Hint "Ouvrir", maintenant qu'il y a un bouton supprimer
  SetControlProperty('BOUVRIR' , 'HINT', TraduireMemoire('Suppression des rubriques sélectionnées'));}
  {JP 09/03/04 : En tréosrerie il est possible de créer des rubriques depuis le Mul (FQ Tréso 10023)}
  SetControlVisible('BINSERT' , True);
  Fiche.Binsert.OnClick := BInsertOnClick;
  {JP 03/11/06 : FQ TRESO 10275 : Ajout de la fonction de contrôle du paramétrage}
  SetControlVisible('BCONTROLRUB', True);
  TToolbarButton97(GetControl('BCONTROLRUB')).OnClick := BControlOnClick;

  {19/07/04 : J'utilise ce boolean pour stocker si les écritures sont soumises
              au calcul des commissions de mouvements}
  SetControlVisible('RB_CATEGORIE', True);
  SetControlCaption('RB_CATEGORIE', TraduireMemoire('Soumis à commission de mouvements'));

  {19/07/04 : au Passage, j'en profite pour afficher le signe}
  RB_SIGNERUB.Values.Clear;
  RB_SIGNERUB.Values.Add('');
  RB_SIGNERUB.Values.Add('C');
  RB_SIGNERUB.Values.Add('D');
  RB_SIGNERUB.Items.Clear;
  {JP 27/12/05 : FQ 10308 : Affichage de Tous : j'ai l'impression que s'il n'y a pas de
                 tablette, il faut le faire manuellement}
  RB_SIGNERUB.Items.Add('<< Tous >>');
  RB_SIGNERUB.Items.Add('Crédit');
  RB_SIGNERUB.Items.Add('Débit');

  {JP 03/10/05 : FQ TRESO 10265}
  SetControlVisible('BDELETE', True);
  SetControlVisible('BDELETE1', False);
  {$ELSE}
  {JP 03/10/05 : FQ TRESO 10265}
  SetControlVisible('BDELETE1', True);
  SetControlVisible('BDELETE', False);
  {$ENDIF}

  {JP 09/09/2003 : Remarque pertinente de Laroche, c'est mieux que dans le OnLoad}
  ChargeCbTypRub ;

  {$IFDEF TRESO}
  {Affectation de la liste ad-hoc}
  TFMul(Ecran).SetDBListe('TRRUBRIQUETRESO');
  {JP 03/10/2005 : FQ TRESO 10264}
  Ecran.Caption := TraduireMemoire('Liste des rubriques');
  UpdateCaption(Ecran);
  {$ENDIF}

  {JP 08/01/08 : FQ 22146 : Mise en place du contrôle du Champ RB_NODOSSIER}
  SetControlVisible('BMAJRUB', V_PGI.SAV);
  if Assigned(GetControl('BMAJRUB')) then
    TToolbarButton97(GetControl('BMAJRUB')).OnClick := BMajRubClick;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPSUPPRUB.OnClose ;
{---------------------------------------------------------------------------------------}
begin
  TDelgene.Clear;
  TNotDel .Clear;
  FreeAndNil(TDelgene);
  FreeAndNil(TNotDel);
  FermeMsg;
  if Ecran.Parent is THPanel then
    _DeblocageMonoPoste(False) ;
  Inherited ;
end ;

{Affichage de la fiche rubrique
{---------------------------------------------------------------------------------------}
procedure TOF_CPSUPPRUB.FListeDblClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var Q : THQuery;
begin
  {$IFDEF EAGLCLIENT}
  if TFMul(Ecran).FListe.RowCount = 0 then Exit;
  Q := TFMul(Ecran).Q;
  Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
  {$ELSE}
  if TFMul(Ecran).Q.Eof and TFMul(Ecran).Q.Bof then Exit;
  Q := TFMul(Ecran).Q;
  {$ENDIF}
  if TFMul(Ecran).FListe.nbSelected > 1 then Exit;

  inherited;


{$IFDEF TRESO}
  if assigned(Q.FindField('RB_NODOSSIER' )) then
    ParametrageRubrique(Q.FindField('RB_RUBRIQUE' ).AsString+';'+Q.FindField('RB_NODOSSIER' ).AsString+';'+Q.FindField('RB_PREDEFINI' ).AsString+';'+Q.FindField('RB_DATEVALIDITE' ).AsString ,taConsult, FCtxSurRubrique)
  else  ParametrageRubrique(Q.FindField('RB_RUBRIQUE' ).AsString,taModif, FCtxSurRubrique);
{$ELSE}
  if assigned(Q.FindField('RB_NODOSSIER' )) then
    ParametrageRubrique(Q.FindField('RB_RUBRIQUE' ).AsString+';'+Q.FindField('RB_NODOSSIER' ).AsString+';'+Q.FindField('RB_PREDEFINI' ).AsString+';'+Q.FindField('RB_DATEVALIDITE' ).AsString ,taConsult, FCtxSurRubrique)
  else  ParametrageRubrique(Q.FindField('RB_RUBRIQUE' ).AsString,taConsult, FCtxSurRubrique);
{$ENDIF}

end;

{---------------------------------------------------------------------------------------}
Procedure TOF_CPSUPPRUB.ChargeCbTypRub ;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRESO}

  {$ELSE}
  { b md 25/07/2005 FQ 16158 }
  if FCtxSurRubrique = CtxBudget then      
  { e md }
  begin
     RB_TYPERUB.DataType  := 'ttRubTypeBud' ;
     TRB_SIGNERUB.Caption := LMsg.Mess[7] ;
     RB_SIGNERUB.Datatype := 'ttBudJal' ;
     RB_SIGNERUB.Name     := 'RB_BUDJAL' ;
   end
   else
     {Récupère les types de rubriques depuis COMMUN}
     FactoriseComboTypeRub(RB_TYPERUB) ;
  {$ENDIF}
end ;


{---------------------------------------------------------------------------------------}
procedure TOF_CPSUPPRUB.BOuvrirClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  i : Integer ;
  j : Byte ;
  {$IFDEF TRESO}
  T : TOB;
  {$ELSE}
  X, Y : DelInfo ;
  {$ENDIF}
  Code, Lib, Typ, Predefini, NoDossier : String ;
  DateValidite : TDateTime;
begin
  NbLig := Fliste.NbSelected ;
  TDelGene.Clear ;
  TNotDel .Clear ;
  {$IFDEF TRESO}
  DateValidite := iDate2099;
  {$ENDIF}
  if NbLig <= 0 then begin
    AfficheMsg(1,'','') ;
    Exit ;
  end ;

  if AfficheMsg(0,'','') <> mrYes then Exit ;

  Effacer    := False ;
  NotEffacer := False ;

  if NbLig > 1 then begin
    {$IFDEF TRESO}
    {10/06/04 : Utilisation d'une tob pour éviter un recours massif aux requêtes}
    T := TOB.Create('***', nil, -1);
    try
      T.LoadDetailFromSQL('SELECT DISTINCT(TE_CODEFLUX) AS CODE FROM TRECRITURE');
    {$ENDIF}
      for i := 0 to NbLig-1 do begin
        Fliste.GotoLeBookMark(i) ;
        {JP 24/02/04 : Je profite du débuggage du else ci-dessous pour ajuster le code}
        Code := GetField('RB_RUBRIQUE');
        Lib  := GetField('RB_LIBELLE');
        Typ  := GetField('RB_TYPERUB');
        {$IFDEF TRESO}
        if T.FindFirst(['CODE'], [Code], True) <> nil then begin
          HShowMessage('1;' + Ecran.Caption + ';La rubrique "' + Code + '" (' + Lib + ') est utilisée dans la trésorerie et ne peut être supprimée.;W;O;O;O', '', '');
          Continue;
        end;
        Detruit(Code, Typ, Predefini, NoDossier, DateValidite);
        {$ELSE}
        Predefini := GetField('RB_PREDEFINI');
        NoDossier := GetField('RB_NODOSSIER');
        DateValidite := GetField('RB_DATEVALIDITE');
        j := Detruit(Code,Typ,Predefini, NoDossier,DateValidite) ;
        {$ENDIF}

        {$IFNDEF TRESO}
        if j <= 0 then begin
          X := DelInfo.Create ;
          X.LeCod := Code ;
          X.LeLib := Lib ;
          X.LeMess:= LMsg.Mess[3] ;
          TDelGene.Add(X) ;
          Effacer := True ;
        end
        else begin
          Y := DelInfo.Create ;
          Y.LeCod := Code ;
          Y.LeLib := Lib ;
          Y.LeMess:= LMsg.Mess[j] ;
          TNotDel.Add(Y) ;
          NotEffacer := True ;
        end;
        {$ENDIF}

      end ;{for}
    {$IFDEF TRESO}
    finally
      FreeAndNil(T);
    end;
    {$ENDIF}
  end

  else begin
    Fliste.GotoLeBookMark(0) ;
    Code := GetField('RB_RUBRIQUE');
    Typ  := GetField('RB_TYPERUB');
    {$IFDEF TRESO}
    if ExisteSQL('SELECT TE_CODEFLUX FROM TRECRITURE WHERE TE_CODEFLUX = "' + Code + '"') then
      PGIError(TraduireMemoire('La rubrique est utilisée dans des écritures de trésorerie.'#13 +
                               'La suppression est impossible.'))
    else begin
    {$ELSE}
    Predefini := GetField('RB_PREDEFINI');
    NoDossier := GetField('RB_NODOSSIER');
    DateValidite := GetField('RB_DATEVALIDITE');
    {$ENDIF}
      j    := Detruit(Code,Typ,Predefini, NoDossier, DateValidite) ;
      if j = 7 then AfficheMsg(8,'','');
    {$IFDEF TRESO}
    end;
    {$ENDIF}
  end ;

  {$IFNDEF TRESO}
  if Effacer    and (AfficheMsg(2,'','') = mrYes) then RapportDeSuppression(TDelGene,1) ;
  if NotEffacer and (AfficheMsg(9,'','') = mrYes) then RapportDeSuppression(TNotDel,1) ;
  {$ENDIF}
  Fiche.BChercheClick(Nil) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPSUPPRUB.Degage ;
{---------------------------------------------------------------------------------------}
var stWhere : string;
begin
  stWhere := 'WHERE RB_RUBRIQUE="' + GeneCode + '" And RB_TYPERUB="'+TypCod+'"';
  {$IFNDEF TRESO}
  stWhere := stWhere + ' AND RB_PREDEFINI="'+fPredefini+'" ';
  stWhere := stWhere + ' AND RB_NODOSSIER="'+fNoDossier+'" ';
  stWhere := stWhere + ' AND RB_DATEVALIDITE="'+USDateTime(fDateValidite)+'" ';
  {$ENDIF TRESO}

  if ExecuteSQL('DELETE FROM RUBRIQUE '+stWhere) <> 1 then
    V_PGI.IoError := oeUnknown ;
end ;

{---------------------------------------------------------------------------------------}
function TOF_CPSUPPRUB.Detruit(St,St1, stPredefini, stNoDossier : string; DateValidite : TDateTime) : Byte ;
{---------------------------------------------------------------------------------------}
begin
  Result   :=0 ;
  GeneCode := St;
  TypCod   := St1;
  fPredefini := stPredefini;
  fNoDossier := stNoDossier;
  fDateValidite := DateValidite;
  if Transactions(Degage, 5) <> oeOK then begin
    MessageAlerte(LMsg.Mess[7]) ;
    Result:=7 ;
  end ;
end ;

{Gestion des message}
{---------------------------------------------------------------------------------------}
procedure TOF_CPSUPPRUB.InitMsg;
{---------------------------------------------------------------------------------------}
begin
  LMsg := THMsgBox.create(FMenuG);
  {00}LMsg.Mess.Add('0;Suppression des rubriques;Confirmez-vous la suppression des enregistrements sélectionnés?;Q;YNC;N;C;');
  {01}LMsg.Mess.Add('1;Suppression des rubriques;Vous n''avez rien sélectionné;E;O;O;O;');
  {02}LMsg.Mess.Add('2;Suppression des rubriques;Désirez-vous un compte-rendu des comptes détruits?;Q;YNC;N;C;');
  {03}LMsg.Mess.Add('Rubrique supprimée');
  {04}LMsg.Mess.Add('Aucun');
  {05}LMsg.Mess.Add('élément sélectionné');
  {06}LMsg.Mess.Add('éléments sélectionnés');
  {07}LMsg.Mess.Add('Compte en cours d''utilisation !');
  {08}LMsg.Mess.Add('8;Suppression des rubriques;Vous ne pouvez pas supprimer ce compte : il est en cours d''utilisation.;W;O;O;O;');
  {09}LMsg.Mess.Add('9;Suppression des rubriques;Désirez-vous un compte-rendu des comptes non détruits?;Q;YNC;N;C;');
  {10}LMsg.Mess.Add('&Budget');
end;

{---------------------------------------------------------------------------------------}
function TOF_CPSUPPRUB.AfficheMsg(num : integer;Av,Ap : string ) : Word ;
{---------------------------------------------------------------------------------------}
begin
  Result := mrNone ;
  if Num > 8 then Exit ;
  Result := LMsg.Execute(num,Av,Ap) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPSUPPRUB.FermeMsg ;
{---------------------------------------------------------------------------------------}
begin
  LMsg.Free ;
end;

{JP 09/03/04 : En Trésorerie, on veut pouvoir créer des rubriques}
{$IFDEF TRESO}
{---------------------------------------------------------------------------------------}
procedure TOF_CPSUPPRUB.BInsertOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
{$IFDEF EAGLCLIENT}
  ParametrageRubrique('', taCreat, FCtxSurRubrique);
{$ELSE}
  ParametrageRubrique('', taCreatEnSerie, FCtxSurRubrique);
{$ENDIF}
  Fiche.BCherche.Click;
end;

{JP 06/11/06 : FQ TRESO 10275 : Contrôle du paramétrage des Rubriques
{---------------------------------------------------------------------------------------}
procedure TOF_CPSUPPRUB.BControlOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  O : TOB;
  T : TOB;
  F : TOB;
  TErr : TOB;
  TTmp : TOB;
  Ok   : Boolean;


      {---------------------------------------------------------------------------------------}
      procedure _TesteCompte(Gen : string);
      {---------------------------------------------------------------------------------------}
      var
        p   : Integer;
        pGen, pGenDet,
        pEx , pExDet : string;
      begin
        for p := 0 to O.Detail.Count - 1 do begin
          F := O.Detail[p];
          pGen := F.GetString('RB_COMPTE1');
          pEx  := F.GetString('RB_EXCLUSION1');
          pGenDet := ReadTokenSt(pGen);
          pExDet  := ReadTokenSt(pEx );
          while pGenDet <> '' do begin
            {On regarde si le compte général fait partie de la rubrique}
            if CompareGeneral(Gen, pGenDet, pExDet) then begin
              {Le compte est concerné par la rubrique courante, on le mémorise}
              with TOB.Create('====', TTmp, -1) do begin
                AddChampSupValeur('COMPTE'  , Gen);
                AddChampSupValeur('RUBRIQUE', F.GetString('RB_COMPTE1'));
                AddChampSupValeur('SENS'    , F.GetString('RB_SIGNERUB'));
                AddChampSupValeur('PAIE'    , F.GetString('RB_EXCLUSION2'));
              end;
            end;
            pGenDet := ReadTokenSt(pGen);
            pExDet  := ReadTokenSt(pEx );
          end;
        end;
      end;


begin
  Ok := HShowMessage('1;' + Ecran.Caption + ';Souhaitez-vous un contrôle complet avec affichage des possibles incohérences'#13 +
                     'sur les modes de paiement (Remarque : les modes de paiement ne sont pas contrôlés).;Q;YN;Y;N', '', '') = mrYes;
  O := TOB.Create('µRUBRIQUE', nil, -1);
  O.LoadDetailFromSQL('SELECT RB_RUBRIQUE, RB_LIBELLE, RB_COMPTE1, RB_EXCLUSION1, RB_EXCLUSION2, RB_SIGNERUB ' +
                      'FROM RUBRIQUE WHERE RB_NATRUB = "TRE" AND RB_CLASSERUB = "TRE"');
  Q := OpenSQL('SELECT G_GENERAL FROM GENERAUX ORDER BY G_GENERAL', True);
  TErr := TOB.Create('µErreur', nil, -1);
  TTmp := TOB.Create('µTEMP', nil, -1);
  InitMove(Q.RecordCount, '');
  try
    while not Q.Eof do begin
      {Recherche du nombre de rubriques concerant ce compte}
      _TesteCompte(Q.FindField('G_GENERAL').AsString);

      {Maintenant on analyse le contenu de la tob Temporaire : }
      F := TTmp.FindFirst(['SENS', 'PAIE'], ['C', ''], True);
      if not Assigned(F) then
        F := TTmp.FindFirst(['SENS'], ['C'], True);
      T := TTmp.FindFirst(['SENS', 'PAIE'], ['D', ''], True);
      if not Assigned(T) then
        T := TTmp.FindFirst(['SENS'], ['D'], True);

      {On n'a pas au moins une ligne de Débit et une ligne de Crédit}
      if (F = nil) or (T = nil) then begin
        with TOB.Create('====', TErr, -1) do begin
          AddChampSupValeur('COMPTE', Q.FindField('G_GENERAL').AsString);
          if not Assigned(F) and not Assigned(T) then AddChampSupValeur('SENS'  , TraduireMemoire('Crédit / Débit non trouvés'))
          else if not Assigned(F) then AddChampSupValeur('SENS'  , TraduireMemoire('Crédit non trouvé'))
          else if not Assigned(T) then AddChampSupValeur('SENS'  , TraduireMemoire('Debit non trouvé'));
          AddChampSupValeur('MOTIF' , TraduireMemoire('Le paramétrage du compte est incomplet'));
        end;
      end
      {Maintenant on précise si les mode de paiement peuvent éventuellement créer des problèmes,
       si l'on a demandé un contrôle complet}
      else if Ok and ((T.GetString('PAIE') <> '') or (F.GetString('PAIE') <> '')) then begin
        with TOB.Create('====', TErr, -1) do begin
          AddChampSupValeur('COMPTE', Q.FindField('G_GENERAL').AsString);
          AddChampSupValeur('SENS', '');
          AddChampSupValeur('MOTIF' , TraduireMemoire('Il y a peut-être des incohérences sur les modes de paiement'));
        end;
      end;
      MoveCur(False);
      Q.Next;
      TTmp.ClearDetail;
    end;

    {Affichage des problèmes}
    if TErr.Detail.Count > 0 then begin
      TheTob := TErr;
      TRAfficheGrille(TraduireMemoire('Contrôle du paramétrage des rubriques'), ag_CtrlRub);
    end
    else
      HShowMessage('0;' + Ecran.caption + ';Le paramétrage est complet;I;O;O;O;', '', '');

  finally
    Ferme(Q);
    if Assigned(TErr) then FreeAndNil(TErr);
    if Assigned(TTmp) then FreeAndNil(TTmp);
    if Assigned(O) then FreeAndNil(O);
    FiniMove;
  end;
end;

{$ENDIF}

{JP 08/01/08 : FQ 22146 : Mise en place du contrôle du Champ RB_NODOSSIER
{---------------------------------------------------------------------------------------}
procedure TOF_CPSUPPRUB.BMajRubClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  CPLanceFiche_CPVERIFRUBDOSSIER('');
end;

initialization
  registerclasses ( [ TOF_CPSUPPRUB ] ) ;
end.
