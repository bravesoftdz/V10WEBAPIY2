{-------------------------------------------------------------------------------------
    Version   | Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
               22/08/01  LG   Création de l'unité
 8.01.001.004  05/02/07  JP   Refonte de l'unité : Seule la partie de sauvegarde du fichier
                              en base (table CETEBAC) est concervée ici. La partie d'intégration
                              dans la table CRELBQE pour la préparation de la saisie de Trésorerie
                              a été déplacé dans la mul sur CETEBAC (CPMULCETEBAC_TOF.PAS)
 8.01.001.012  24/04/07  JP   FQ TRESO 10441 : Gestion du Is Null pour Oracle sur CET_IMO
 8.01.001.015  14/05/07  JP   FQ TRESO 10464 : Ajout du Bouton de visualisation
 8.01.001.020  19/06/07  JP   FQ TRESO 10476 : Gestion du bouton Chercher
 8.10.001.013  08/10/07  JP   Ajout des soldes dans la grilles, pour permettre un contrôle visuel de leur cohérence
 8.10.002.001  16/10/07  JP   FQ 21656 : autorisation de l'accès en compta si pointage sur TRECRITURE => ajout d'un msg
 8.10.002.001  22/10/07  JP   FQ 21706 : Gestion des fichiers sans retour de chariot
 8.10.005.001  14/11/07  JP   bouton BDELETE ne sert à rien, on le cache
--------------------------------------------------------------------------------------}
unit dpTOFCETEBAC;

interface

uses
  StdCtrls, Controls, Classes, SysUtils,
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  ComCtrls, HCtrls, HEnt1, UTOF, UTOB,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main,
  {$ENDIF}
  CPProcGen,
  Vierge, HTB97,uEntCommun;


type
  TOF_CETEBAC = class (TOF)
    HGFichier : THGrid;
    BValider  : TToolbarButton97; {Bouton Valider : coche verte}
    BTag      : TToolbarButton97; {Bouton tout selectionner}
    BDeTag    : TToolbarButton97; {Bouton tout deselectionner}
    cbBanque  : THValComboBox;

    procedure BValiderClickFichier(Sender : TObject);
    procedure BDeTagClickFichier( Sender : TObject);
    procedure BTagClickFichier( Sender : TObject);
    procedure HGFichierDblClick   (Sender : TObject);
    procedure FormKeyDown         (Sender : TObject; var Key : Word; Shift : TShiftState);
    procedure ImprimerClick       (Sender : TObject);
    procedure RepertoireExit      (Sender : TObject);
    procedure CompteBqeChange     (Sender : TObject);

    procedure GetControlTOF;
    function  AssignEvent  : Boolean;
  private
    NumLigne : Integer;
    FDecimal : Integer;

    procedure GereRepertoire    ;
    procedure InitGrille        ;
    procedure RemplitGrid       ;
    procedure ControleFichiers  (Fic : string);{JP 23/10/07 : FQ 21706 : Gestion des fichiers Multi-relevés ou/et sans retour de chariot}
    procedure RemplitTobImp     (Fic : string; var T : TOB);
    procedure RempliLigne       (TR : TOB; StReleve : string; NumReleve : Integer);
    procedure RempliSoldeInitial(TR : TOB; StReleve : string; NumReleve : Integer);
    procedure RempliSoldeFinal  (TR : TOB; StReleve : string; NumReleve : Integer);
    function  ControleReleve    (var TobRel : TOB) : Boolean;
    procedure SauvegardeFichier (NomFichier : string; Fichier : TStringList; Del : Boolean);
    procedure RempliLibelleComplementaire(TR, TRLigne : TOB; StReleve : string);
  public
    procedure OnArgument (S : String ) ; override ;
    function  OnIntegreReleve(NomReleve : string; Efface : Boolean; var TobErr : Tob) : Boolean;
  end;


procedure CPLanceFiche_ImportTresorerie ( vListeErreur : TStringList ) ;
procedure CPLanceFiche_ImportCEtebac(Arg : string);

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  Dialogs, ParamSoc, Windows,
  {$IFDEF COMPTA}
  ImRapInt,
  {$ENDIF}
  {$IFDEF EAGLCLIENT}
  UtileAgl,
  {$ELSE}
  EdtREtat, StdConvs,
  {$ENDIF EAGLCLIENT}
  HMsgBox, EtbUser, Ent1, Commun, Constantes, UObjEtats, dpTOFVisuReleve,
  Grids, cbpPath, ULibPointage;

function ConvertDate(DateChar : string) : TDateTime; forward;
function ConvertMontant(MontantChar : string; Deci : Integer) : Double; forward;

const
 {Grid Releve}
 cColNomFichier  = 0;
 cColRib         = 1;
 cColDu          = 2;
 cColAu          = 3;
 cColDevise      = 4;
 cColSoldeDeb    = 5; {JP 08/10/07}
 cColSoldeFin    = 6; {JP 08/10/07}

{---------------------------------------------------------------------------------------}
procedure CPLanceFiche_ImportCEtebac(Arg : string);
{---------------------------------------------------------------------------------------}
begin
  {JP 16/10/07 : FQ 21656 : ajout d'un message si on est en compta mais pointage sur TRECRITURE}
  if (CtxCompta in V_PGI.PGIContexte) and EstPointageCache then begin
    if PgiAsk(TraduireMemoire('Vous êtes en mode "Pointage dans la Trésorerie". L''intégration de relevés ') + #13 +
              TraduireMemoire('en Comptabilité ne doit être utilisée que pour la seule saisie de trésorerie.') + #13#13 +
              TraduireMemoire('Souhaitez-vous abandonner ?'), TraduireMemoire('Intégration de relevés')) = mrNo then
      AGLLanceFiche('CP','RLVETEBAC','','', Arg);
  end
  else
    AGLLanceFiche('CP','RLVETEBAC','','', Arg);
end;

{---------------------------------------------------------------------------------------}
procedure CPLanceFiche_ImportTresorerie(vListeErreur : TStringList);
{---------------------------------------------------------------------------------------}

    {------------------------------------------------------------------}
    function _GetListeFichier : TStringList;
    {------------------------------------------------------------------}
    var
     MySearchRec   : TSearchRec;
     lStNomReleve  : string;
     lStRepReleve  : string ;
     lBoOK         : Boolean;
    begin

      Result := TStringList.Create ;

      if (ctxPCL in V_PGI.PGIContexte) then
        lStRepReleve := GetParamSocSecur ('SO_RETEBAC4', TcbpPath.GetCegidDistriApp + '\ETEBAC')
      else
        lStRepReleve := GetParamSocSecur ('SO_REPETEBAC3', TcbpPath.GetCegidDistriApp + '\ETEBAC') ;

      if length( lStRepReleve ) = 0 then begin
        if ( ctxPCL in V_PGI.PGIContexte ) then
          PGIInfo(TraduireMemoire('Répertoire Etebac non définie (menu Etebac-Paramètrage Répertoire)'), TraduireMemoire('Erreur de paramètrage'))
        else
          PGIInfo(TraduireMemoire('Répertoire Etebac non définie ( Menu Suivi tiers - Répertoire etebac )'), TraduireMemoire('Erreur de paramètrage'));
        Exit;
      end;

     if lStRepReleve[length(lStRepReleve)] <> '\' then
      lStRepReleve := lStRepReleve + '\' ;

     lStRepReleve := lStRepReleve + '*.*';

     lBoOK := FindFirst(lStRepReleve, faAnyFile, MySearchRec) = 0 ;

     try

      if lBoOK then
       repeat
        if not (MySearchRec.Attr and faDirectory > 0) then  // On ne prend pas les répertoires. - CA - 29/09/2003
        begin
          lStNomReleve := ExtractFilePath(lStRepReleve) + MySearchRec.Name ;
          if ( pos('~',lStNomReleve) > 0 ) or ( pos('.INI',lStNomReleve) > 0 ) then continue ;
          result.Add (lStNomReleve );
        end;
       until FindNext(MySearchRec)<>0 ;


     finally
      SysUtils.FindClose(MySearchRec) ;
     end ;
    end;

var
 lMyTOF : TOF_CETEBAC;
 lListe : TStringList;
 i      : Integer;
 TobErr : TOB;
begin

  lMyTOF  := TOF_CETEBAC.Create(nil,false) ;
  lListe  := _GetListeFichier ;
  TobErr  := TOB.Create('***', nil, -1);
  try
    for  i := 0 to lListe.count - 1 do begin
      lMyTOF.OnIntegreReleve(lListe[i], False, TobErr);
      {$IFNDEF TT}
      SysUtils.DeleteFile(lListe[i]);
      {$ENDIF}
     end ; // for
  finally
    FreeAndNil(TobErr);
  end;

  lMyTOF.Free;
  lListe.Free;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CETEBAC.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 7773000;
  GetControlTOF;
  if not AssignEvent   then exit;
  {Gestion du répertoire}
  GereRepertoire;
  {Remplissage de la grille}
  RemplitGrid;
  {Initialisation éventuelle du compte}
  if S <> '' then
    cbBanque.Value := ReadTokenSt(S);

  BDeTag.Visible := false ;
  BDEtag.Left    := BTag.Left ;
  {JP 14/11/07 : FQ BUREAU 11813 : Le bouton ne sert à rien}
  SetControlVisible('BDELETE', False);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CETEBAC.GetControlTOF;
{---------------------------------------------------------------------------------------}
begin
 HGFichier := THGrid(GetControl('HGFICHIER', True));
 BValider  := TToolbarButton97(GetControl('BVALIDER', True));
 BTag      := TToolbarButton97(GetControl('BTAG'    , True));
 BDeTag    := TToolbarButton97(GetControl('BDETAG'  , True));
 cbBanque  := THValComboBox(GetControl('CBANQUE'    , True));
end;


{---------------------------------------------------------------------------------------}
function TOF_CETEBAC.AssignEvent : boolean;
{---------------------------------------------------------------------------------------}
begin
  BValider.OnClick     := BValiderClickFichier;
  Ecran.OnKeyDown      := FormKeyDown;
  BTag.OnClick         := BTagClickFichier;
  BDeTag.OnClick       := BDeTagClickFichier;
 // BDeTag.OnClick       := BTagClickFichier;
  HGFichier.OnDblClick := HGFichierDblClick;

  {14/05/07 : FQ TRESO 10464 : ajout du Bouton de visualisation}
  if Assigned(GetControl('BVOIRRELEVE')) then
    (GetControl('BVOIRRELEVE') as TToolbarButton97).OnClick := HGFichierDblClick;

  {JP 21/09/04 : FQ TRESO 10120 : possibilité d'imprimer plusieurs relevés}
  TToolbarButton97(GetControl('BIMPRIMER')).OnClick := ImprimerClick;
  cbBanque.OnChange := CompteBqeChange;

  {JP 19/06/07 : FQ TRESO 10476 : gestion du bouton rechercher}
  if Assigned(GetControl('BCHERCHER')) then begin
    SetControlVisible('BCHERCHER', True);
    TToolbarButton97(GetControl('BCHERCHER')).OnClick := RepertoireExit;
  end;

  Result := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CETEBAC.FormKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin
  case Key of
    VK_F10   : BValiderClickFichier(BValider); {JP 19/07/07}
    Ord('A') : if (Shift = [ssCtrl]) then begin
                 BTagClickFichier(nil);
                 Key := 0;
               end;
    Ord('D') : if (Shift = [ssAlt]) then HGFichierDblClick(nil);
    else
      TFVierge(Ecran).FormKeyDown(Sender, Key, Shift);
  end;
end;

{Visualisation du relevé en cours
{---------------------------------------------------------------------------------------}
procedure TOF_CETEBAC.HGFichierDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
 lStArgument : string;
begin

  lStArgument := 'NOMFICHIER=' + HGFichier.Cells[cColNomFichier, HGFichier.Row] +
                 {Comme on ne peut afficher qu'un seul fichier, il faut mettre 1
                 ';NUMFICHIER=' +IntToStr(HGFichier.Row) ;}
                 ';NUMFICHIER=1';

  {Ouverture de la fenetre de visu des releves}
  CPLanceFicheVisuRel(lStArgument);
end;


{
procedure TOF_CETEBAC.BTagClickFichier(Sender : TObject);
begin
  if HGFichier.AllSelected then
   HGFichier.ClearSelected
    else
     HGFichier.AllSelected := True;
end; }

procedure TOF_CETEBAC.BTagClickFichier( Sender : TObject);
begin
 HGFichier.AllSelected := true;
 BDeTag.Visible        := true ;
 BTag.Visible          := false ;
end;

procedure TOF_CETEBAC.BDeTagClickFichier( Sender : TObject);
begin
 HGFichier.ClearSelected;
 BDeTag.Visible := false ;
 BTag.Visible   := true ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CETEBAC.GereRepertoire;
{---------------------------------------------------------------------------------------}
begin
  if (ctxTreso in V_PGI.PGIContexte) then begin
    cbBanque.DataType := 'TRBANQUECP';
    cbBanque.Plus := FiltreBanqueCp('TRBANQUECP', tcb_Bancaire, '');
  end
  else begin
    cbBanque.DataType := 'TTBANQUECP';
    SetPlusBanqueCp(cbBanque);
  end;

  {En PCL, L'utilisateur ne peut choisir le répertoire
  SetControlenabled('REPERTOIRE' , not (ctxPCL in V_PGI.PGIContexte));
  SetControlEnabled('REPERTLABEL', not (ctxPCL in V_PGI.PGIContexte));}
  SetControlVisible('PNTOP' , not (ctxLanceur in V_PGI.PGIContexte));

  if (ctxPCL in V_PGI.PGIContexte) then
    SetControlText('REPERTOIRE', GetParamSocSecur('SO_RETEBAC4', TcbpPath.GetCegidDistriApp + '\ETEBAC'))
  else begin
    SetControlText('REPERTOIRE', GetParamSocSecur('SO_REPETEBAC3', TcbpPath.GetCegidDistriApp + '\ETEBAC'));
    (GetControl('REPERTOIRE') as THEdit).OnExit := RepertoireExit;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CETEBAC.RepertoireExit(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  CompteBqeChange(cbBanque);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CETEBAC.CompteBqeChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  RemplitGrid;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CETEBAC.InitGrille;
{---------------------------------------------------------------------------------------}
var
  Largeur : Integer;
begin
  HGFichier.RowCount  := 2;
  HGFichier.FixedRows := 1;
  HGFichier.ColCount  := 7;{JP 08/10/07 : ajout de deux colonnes}
  HGFichier.Cells[cColNomFichier, 0] := TraduireMemoire('Fichier');
  HGFichier.Cells[cColRib       , 0] := TraduireMemoire('RIB');
  HGFichier.Cells[cColDu        , 0] := TraduireMemoire('Du');
  HGFichier.Cells[cColAu        , 0] := TraduireMemoire('Au');
  HGFichier.Cells[cColDevise    , 0] := TraduireMemoire('Devise');
  HGFichier.Cells[cColSoldeDeb  , 0] := TraduireMemoire('Solde deb.');{JP 08/10/07}
  HGFichier.Cells[cColSoldeFin  , 0] := TraduireMemoire('Solde fin');{JP 08/10/07}

  HGFichier.FColAligns[cColRib       ] := taCenter;
  HGFichier.FColAligns[cColDu        ] := taCenter ;
  HGFichier.FColAligns[cColAu        ] := taCenter;
  HGFichier.FColAligns[cColDevise    ] := taCenter;
  HGFichier.FColAligns[cColSoldeDeb  ] := taRightJustify;{JP 08/10/07}
  HGFichier.FColAligns[cColSoldeFin  ] := taRightJustify;{JP 08/10/07}

  Largeur := HGFichier.Width - 25;
  HGFichier.ColWidths[cColNomFichier] := Round(Largeur * 0.30);
  HGFichier.ColWidths[cColRib       ] := Round(Largeur * 0.20);
  HGFichier.ColWidths[cColDu        ] := Round(Largeur * 0.10);
  HGFichier.ColWidths[cColAu        ] := Round(Largeur * 0.10);
  HGFichier.ColWidths[cColDevise    ] := Round(Largeur * 0.08);
  HGFichier.ColWidths[cColSoldeDeb  ] := Round(Largeur * 0.11);{JP 08/10/07}
  HGFichier.ColWidths[cColSoldeFin  ] := Round(Largeur * 0.11);{JP 08/10/07}

  {Pour que les tri sur colonne soient numérique et non alpha-nuémrique}
  HGFichier.SortEnabled := True;
  HGFichier.ColTypes[cColDu] := 'D';
  HGFichier.ColTypes[cColAu] := 'D';
  {JP 08/10/07 : Paramétrage des colonnes de soldes}
  HGFichier.ColTypes[cColSoldeDeb] := 'R';
  HGFichier.ColTypes[cColSoldeFin] := 'R';
  HGFichier.FColFormats[cColSoldeDeb] := StrFMask(V_PGI.OkDecV, '', True);
  HGFichier.FColFormats[cColSoldeFin] := StrFMask(V_PGI.OkDecV, '', True);
end;

{JP 23/10/07 : FQ 21706 : Gestion des fichiers Multi-relevés ou/et sans retour de chariot
{---------------------------------------------------------------------------------------}
procedure TOF_CETEBAC.ControleFichiers(Fic : string);
{---------------------------------------------------------------------------------------}
var
  MySearchRec  : TSearchRec;
  ResultSearch : Integer;
  NomReleve    : string;
begin

  ResultSearch := FindFirst(Fic, faAnyFile, MySearchRec);
  try
    while ResultSearch = 0 do begin

      if (MySearchRec.Name = '.') or (MySearchRec.Name = '..') or (MySearchRec.Attr = faDirectory) or
         (Pos('.INI', MySearchRec.Name) > 0) or (Pos('~', MySearchRec.Name) > 0) then begin
        ResultSearch := FindNext(MySearchRec);
        Continue;
      end;

      NomReleve := ExtractFilePath(Fic) + MySearchRec.Name;
      {Ajout des retours chariot si neccéssaire}
      TransformeFichier(NomReleve);
      {Création d'autant de fichiers que de relevés
       23/10/07 : Non en PCL, en attente de validation de la part de la qualité PGE
                  Autre possibilité : ajouter un bouton qui fera le traitement
      if (ctxPcl in V_PGI.PGIContexte) then DecoupeFichier(NomReleve, '01', '07', 'DAT');}

      ResultSearch := FindNext(MySearchRec);
    end;
  finally
    SysUtils.FindClose(MySearchRec);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CETEBAC.RemplitGrid;
{---------------------------------------------------------------------------------------}
var
  MySearchRec : TSearchRec;
  FSource     : Textfile;
  RepReleve   : string;
  StReleve    : string;
  NomReleve   : string;
  rDevise     : string;
  rBanque     : string;
  rGuichet    : string;
  rNumCompte  : string;
  tDevise     : string;
  tBanque     : string;
  tGuichet    : string;
  tNumCompte  : string;
  QQ : TQuery ;
  Okok: Boolean;
  SQL : string ;
  DateDu : string;
  DateAu : string;
  ch     : string;
  n      : Integer;
begin
  {Prépare la grille}
  InitGrille;

  RepReleve := GetControlText('REPERTOIRE');

  if RepReleve = '' then begin
    PGIError(TraduireMemoire('Le nom du répertoire ETEBAC n''est pas valide'), Ecran.Caption);
    Exit;
  end;

  if RepReleve[Length(RepReleve)] <> '\' then RepReleve := RepReleve + '\';
  if not DirectoryExists(RepReleve) then begin
    PGIError(TraduireMemoire('Le nom du répertoire ETEBAC n''est pas valide'), Ecran.Caption);
    Exit;
  end;

  RepReleve := RepReleve + '*.*';

  for n := 0 to HGFichier.ColCount - 1 do HGFichier.Cells[n, 1] := '';

  {JP 23/10/07 : FQ 21706 : Gestion des fichiers Multi-relevés ou/et sans retour de chariot}
  ControleFichiers(RepReleve);

  Okok := (FindFirst(RepReleve, faAnyFile, MySearchRec) = 0);
  try
    if Okok then begin
      {On commence par récupérer les opérations du compte de sélection}
      if (cbBanque.Value <> '') then begin
        SQL := 'SELECT BQ_ETABBQ, BQ_GUICHET, BQ_NUMEROCOMPTE, D_CODEISO ' +
               'FROM BANQUECP LEFT JOIN DEVISE ON D_DEVISE = BQ_DEVISE ';
        if (ctxTreso in V_PGI.PGIContexte) then SQL := SQL + 'WHERE BQ_CODE = "' + cbBanque.Value + '"'
                                           else SQL := SQL + 'WHERE BQ_GENERAL = "' + cbBanque.Value + '" AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '"';
        QQ := OpenSql(SQL, True,-1,'',true);
        try
          if not QQ.Eof then begin
            tDevise    := QQ.FindField('D_CODEISO').AsString;
            tBanque    := QQ.FindField('BQ_ETABBQ').AsString;
            tGuichet   := QQ.FindField('BQ_GUICHET').AsString;
            tNumCompte := QQ.FindField('BQ_NUMEROCOMPTE').AsString;
          end;
        finally
          ferme(QQ);
        end ;
      end;

      repeat
        if (MySearchRec.Name = '.') or (MySearchRec.Name = '..') or (MySearchRec.Attr = faDirectory) or
           (Pos('.INI', MySearchRec.Name) > 0) or (Pos('~', MySearchRec.Name) > 0) then
          Continue;

        NomReleve := ExtractFilePath(RepReleve) + MySearchRec.Name;

        AssignFile(FSource, NomReleve);
        Reset(FSource);
        try
          if not EOF(FSource) then begin
            Readln(FSource, StReleve);
            if VH^.PaysLocalisation = CodeISOES then begin
              rBanque    := Copy(StReleve, 3, 4);
              rGuichet   := Copy(StReleve, 7, 4);
              rDevise    := Copy(StReleve, 48, 3);
              rNumCompte := Copy(StReleve, 11, 10);
              DateDu    := DateToStr(EtbUser.ConvertDate(Copy(StReleve, 21, 6), 'AAMMDD'));
              DateAu    := DateToStr(EtbUser.ConvertDate(Copy(StReleve, 27, 6), 'AAMMDD'));
            end
            else begin
              rBanque    := Copy(StReleve, 3, 5);
              rGuichet   := Copy(StReleve, 12, 5);
              rDevise    := Copy(StReleve, 17, 3);
              rNumCompte := Copy(StReleve, 22, 11);
              {JP 26/07/05 : FQ 16162 : pour éviter si un relevé n'est pas correct que tous les suivants
                             ne soient pas lus à cause d'un plantage}
              ch := Copy(StReleve, 35, 6);
              if IsNumeric(ch) then DateDu := DateToStr(ConvertDate(ch))
                               else Continue;
            end;

            {Soit il n'y a pas de filtre sur les comptes, soit on s'assure que le fichier en cours
             correspond bien au compte sélectionné}
            OkOk := (cbBanque.Value = '') or
                    ((rBanque = tBanque) and (rGuichet = tGuichet) and (rNumCompte = tNumCompte) and (rDevise = tDevise));

            if Okok then begin
              HGFichier.Cells[cColNomFichier, HGFichier.RowCount - 1] := NomReleve;
              HGFichier.Cells[cColRib       , HGFichier.RowCount - 1] := rBanque + ' ' + rGuichet + ' '+ rNumCompte;
              HGFichier.Cells[cColDu        , HGFichier.RowCount - 1] := DateDu;
              HGFichier.Cells[cColDevise    , HGFichier.RowCount - 1] := rDevise;
              {JP 08/10/07 : récupération du solde initial}
              HGFichier.Cells[cColSoldeDeb  , HGFichier.RowCount - 1] := FloatToStr(ConvertMontant(Copy(StReleve, 91, 14), StrToInt(Copy(StReleve, 20, 1))));
              if VH^.PaysLocalisation <> CodeISOES then begin

                while (Copy(StReleve, 1, 2) <> '07') and (not Eof(FSource)) do Readln(FSource, StReleve);

                if Copy(StReleve, 1, 2) = '07' then begin
                  {JP 08/10/07 : récupération du solde final}
                  HGFichier.Cells[cColSoldeFin  , HGFichier.RowCount - 1] := FloatToStr(ConvertMontant(Copy(StReleve, 91, 14), StrToInt(Copy(StReleve, 20, 1))));

                  {JP 26/07/05 : FQ 16162 : pour éviter si un relevé n'est pas correct que tous les suivants
                                 ne soient pas lus à cause d'un plantage}
                  ch := Copy(StReleve, 35, 6);
                  if IsNumeric(ch) then HGFichier.Cells[cColAu, HGFichier.RowCount - 1] := DateToStr(ConvertDate(ch))
                                   else Continue;
                end
                else
                  HGFichier.Cells[cColau, HGFichier.RowCount - 1] := '  /  /    ';
              end;
              HGFichier.RowCount := HGFichier.RowCount + 1;
            end;
          end;
        finally
          CloseFile(FSource);
        end;
      until FindNext(MySearchRec) <> 0;
    end;

    if HGFichier.RowCount <> 2 then HGFichier.RowCount := HGFichier.RowCount - 1;
  finally
    SysUtils.FindClose(MySearchRec);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CETEBAC.BValiderClickFichier( Sender : TObject);
{---------------------------------------------------------------------------------------}


    {---------------------------------------------------------------}
    procedure _LanceEtatErr(var TobErr : TOB);
    {---------------------------------------------------------------}
    var
      n : Integer;
      T : TOB;
      F : TOB;
      O : TObjEtats;
    begin
      {Préparation de la tob d'impression}
      T := TOB.Create('****', nil, -1);
      try
        for n := 0 to TobErr.Detail.Count - 1 do begin
          F := TOB.Create('µµµ', T, -1);
          F.AddChampSupValeur('ERREUR', TobErr.Detail[n].GetString('ERREUR'));
          F.AddChampSupValeur('LIGNE', TobErr.Detail[n].GetString('CET_TYPELIGNE'));
          F.AddChampSupValeur('DATEOPE', TobErr.Detail[n].GetString('CET_DATEOPERATION'));
          F.AddChampSupValeur('RIB', TobErr.Detail[n].GetString('CET_ETABBQ') + ' ' + TobErr.Detail[n].GetString('CET_GUICHET') + ' ' + TobErr.Detail[n].GetString('CET_NUMEROCOMPTE'));
          F.AddChampSupValeur('LIBELLE', TobErr.Detail[n].GetString('CET_LIBELLE'));
          F.AddChampSupValeur('LIBENRI', TobErr.Detail[n].GetString('CET_LIBELLE1'));
          F.AddChampSupValeur('DEBIT',  TobErr.Detail[n].GetDouble('CET_DEBIT') + TobErr.Detail[n].GetDouble('CET_DEBITDEV'));
          F.AddChampSupValeur('CREDIT', TobErr.Detail[n].GetDouble('CET_CREDIT') + TobErr.Detail[n].GetDouble('CET_CREDITDEV'));
        end;

        TobErr.ClearDetail;
        for n := T.Detail.Count - 1 downto 0 do
          T.Detail[n].ChangeParent(TobErr, 0);
      finally
        FreeAndNil(T);
      end;
      TobErr.Detail.Sort('RIB;LIGNE;DATEOPE;');
      {Paramétrage de l'état}
      O := TObjEtats.Create(TraduireMemoire('Récapitulatif des erreurs d''intégration'), TobErr);
      try
        O.MajTitre(0, TraduireMemoire('Erreur'));
        O.MajTitre(1, TraduireMemoire('Type de ligne'));
        O.MajTitre(2, TraduireMemoire('Date d''opération'));
        O.MajTitre(3, TraduireMemoire('Rib'));
        O.MajTitre(4, TraduireMemoire('Libellé'));
        O.MajTitre(5, TraduireMemoire('Lib. enrichi'));
        O.MajTitre(6, TraduireMemoire('Débit'));
        O.MajTitre(7, TraduireMemoire('Crédit'));

        O.MajAlign(0, ali_Gauche);
        O.MajAlign(1, ali_Centre);
        O.MajAlign(2, ali_Centre);
        O.MajAlign(3, ali_Gauche);
        O.MajAlign(4, ali_Gauche);
        O.MajAlign(5, ali_Gauche);
        O.MajAlign(6, ali_Droite);
        O.MajAlign(7, ali_Droite);

        O.MajFormat(1, 'C');
        O.MajFormat(3, 'C');
        O.MajFormat(6, StrFMask(V_PGI.OkDecV, '', True));
        O.MajFormat(7, StrFMask(V_PGI.OkDecV, '', True));
        O.Imprimer;
      finally
        FreeAndNil(O);
      end;
    end;

var
  i         : Integer;
  lBoEfface : Boolean;
  TobErr    : TOB;
  Resultat  : string;
begin
  if (HGFichier.AllSelected) or (HGFichier.nbSelected <> 0) then begin

    lBoEfface := HShowMessage('0;' + Ecran.Caption + ';' + TraduireMemoire('Voulez-vous supprimer tous les fichiers après intégration.') +
                              ';I;YN;N;N', '', '') = mrYes;
    //PGIAsk(TraduireMemoire('Voulez-vous supprimer tous les fichiers après intégration.'), Ecran.Caption) = mrYes;
    TobErr    := TOB.Create('_ERREUR', nil, - 1);
    {Ce champ contiendra le nombre de fichiers traités}
    TobErr.AddChampSupValeur('NBFICHIER', 0);
    try
      {Lancement du traitement}
      {
      for i := 1 to (HGFichier.RowCount - 1) do begin
        if HGFichier.IsSelected(i) then
        }
      for i := 0 to (HGFichier.nbSelected - 1) do begin
        HGFichier.GotoLeBookMark(i);
        OnIntegreReleve(HGFichier.Cells[cColNomFichier,HGFichier.Row], lBoEfface, TobErr);
      end;

      {Impression du rapport d'erreurs}
      if TobErr.Detail.Count > 0 then
        _LanceEtatErr(TobErr);

      {Affichage du traitement}
      if TobErr.GetInteger('NBFICHIER') = 0 then
        Resultat := TraduireMemoire('Aucun fichier n''a été intégré')
      else if TobErr.GetInteger('NBFICHIER') = 1 then
        Resultat := TraduireMemoire('Un fichier a été intégré')
      else
        Resultat := TobErr.GetString('NBFICHIER') + ' ' + TraduireMemoire('fichiers ont été intégrés');
      PGIInfo(Resultat);

    finally
      FreeAndNil(TobErr);
    end;
  end
  else
    PGIINFO(TraduireMemoire('Vous n''avez pas sélectionné de fichier'), Ecran.Caption);

  {Rafraîchissement de la grille}  
  HGFichier.ClearSelected;
  RemplitGrid;
end;

{Intégration proprement dite dans la table CETEBAC
{---------------------------------------------------------------------------------------}
function TOF_CETEBAC.OnIntegreReleve(NomReleve : string; Efface : Boolean; var TobErr : Tob) : Boolean;
{---------------------------------------------------------------------------------------}

    {---------------------------------------------------------------}
    procedure _DeplaceTob(var Source : TOB; var Dest : TOB);
    {---------------------------------------------------------------}
    var
      n, p : Integer;
    begin
      p := Dest.Detail.Count - 1;
      for n := Source.Detail.Count - 1 downto 0 do
        Source.Detail[n].ChangeParent(Dest, p);
    end;

var
  StReleve, CodeOpe : string;
  TobReleve, TR  : Tob;
  TobTmp         : TOB;
  NumReleve      : Integer;
  LFichier       : TStringList;
  IsFichierOk    : Boolean;
  Q : TQuery;
  n : Integer;
begin
  Result := False;
  NumReleve := 0;

  LFichier  := TStringList.Create;
  TobTmp    := TOB.Create('_TEMPOR', nil, - 1);
  TobReleve := TOB.Create('_RELEVE', nil, - 1);
  try
    LFichier.LoadFromFile(NomReleve);
    TR := nil;
    {Un fichier pouvant contenir plusieurs relevés, ce booléen est passer à false
     dès qu'un relevé n'est pas correct : dans ce cas, on n'insère pas en table et
     on ne détruit pas le fichier, même si Efface est à True}
    IsFichierOk := True;

    {lecture du fichier}
    for n := 0 to LFichier.Count - 1 do begin
      StReleve := LFichier[n];
      CodeOpe := Copy(StReleve, 1, 2);
      {Gestion de la ligne de solde Entete : Ancien Solde}
      if (CodeOpe = '01')  then begin
        if NumReleve = 0 then begin
          Q := OpenSQL('SELECT MAX(CET_NUMRELEVE) AS N FROM CETEBAC', True,-1,'',true);
          try
            NumReleve := Q.FindField('N').AsInteger + 1;
          finally
            Ferme(Q);
          end;
        end
        else
          Inc(NumReleve);

        NumLigne := 0;
        TR := TOB.Create('CETEBAC', TobTmp, -1);
        RempliSoldeInitial(TR, StReleve, NumReleve);
      end;

      {Gestion des lignes d'écritures}
      if (CodeOpe = '04') and (TR <> nil)  then begin
        TR := TOB.Create('CETEBAC', TobTmp, -1);
        RempliLigne(TR, StReleve, NumReleve);
      end;

     {Gestion des lignes de commentaires supplementaires}
      if (CodeOpe = '05') and (TR <> nil) then
        RempliLibelleComplementaire(TobTmp, TR, StReleve);

     {Gestion de la ligne de Nouveau Solde}
      if (CodeOpe = '07') and (TR  <> nil)  then begin
        TR := TOB.Create('CETEBAC', TobTmp, -1);
        RempliSoldeFinal(TR, StReleve, NumReleve);

        {Contrôle de la cohérence du relevé : en mode PCL, renvoie True,
         les contrôles étant faits lors de la création d'une session de pointage}
        if ControleReleve(TobTmp) then
          _DeplaceTob(TobTmp, TobReleve)
        else begin
          {On rajoute le nom du fichier pour le rapport d'erreur : à ce niveau,
           TobTmp a obligatoirement au moins une fille}
          TobTmp.Detail[0].AddChampSupValeur('FICHIER', NomReleve, True);
          {On n'arrête pas le traitement, afin de notifier l'ensemble des erreurs du fichier}
          IsFichierOk := False;
          {On met le relevé dans la tob des erreurs pour le rapport}
          _DeplaceTob(TobTmp, TobErr);
        end;

        {Par précaution !!}
        TobTmp.ClearDetail;
      end;
    end;

    if IsFichierOk and (TobReleve.Detail.Count > 0) then begin
      TobErr.SetDouble('NBFICHIER', TobErr.GetDouble('NBFICHIER') + 1);
      TobReleve.InsertDB(nil, True);
      Result := True;
    end;

    {Je n'efface ou sauvegarde le fichier que s'il n'y a pas d'erreur}
    if IsFichierOk then
      SauvegardeFichier(NomReleve, lFichier, Efface);


  finally
    FreeAndNil(TobReleve);
    FreeAndNil(TobTmp);
    FreeAndNil(LFichier);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CETEBAC.RempliSoldeInitial(TR : TOB; StReleve : string; NumReleve : Integer);
{---------------------------------------------------------------------------------------}
var
  Montant : double;
  Devise : String ;
begin
  Inc(NumLigne);

  TR.PutValue('CET_NUMRELEVE', NumReleve);
  TR.PutValue('CET_ASUPPRIMER', '-');
  TR.PutValue('CET_TYPELIGNE', '01' );
  TR.PutValue('CET_NUMLIGNE', NumLigne);
  TR.PutValue('CET_ETABBQ', Copy(StReleve, 3, 5));
  TR.PutValue('CET_GUICHET', Copy(StReleve, 12, 5));
  TR.PutValue('CET_NUMEROCOMPTE', Copy(StReleve, 22, 11));
  TR.PutValue('CET_CODEAFB', '');
  TR.PutValue('CET_DATEOPERATION', ConvertDate(Copy(StReleve, 35, 6)));

  TR.PutValue('CET_EXONERE', 0);
  TR.PutValue('CET_IMO', '-'); {24/04/07 : FQ TRESO 10441 : pour Oracle}
  TR.PutValue('CET_DATEVALEUR', ConvertDate(Copy(StReleve, 35, 6)));
  TR.PutValue('CET_LIBELLE', '');
  TR.PutValue('CET_LIBELLE1', '');
  TR.PutValue('CET_LIBELLE2', '');
  TR.PutValue('CET_LIBELLE3', '');
  TR.PutValue('CET_NUMEROPIECE', '');
  TR.PutValue('CET_REFPIECE', '');
  TR.PutValue('CET_CONTROLOK', '-');

  Devise := Copy(StReleve, 17, 3);
  TR.PutValue('CET_DEVISE', Devise);

  TR.SetDouble('CET_CREDIT'   , 0);
  TR.SetDouble('CET_DEBIT'    , 0);
  TR.SetDouble('CET_CREDITDEV', 0);
  TR.SetDouble('CET_DEBITDEV' , 0);

  {Ici, il s'agit de l'ancien solde}
  Montant := ConvertMontant(Copy(StReleve, 91, 14), StrToInt(Copy(StReleve, 20, 1)));
  if Montant > 0 then begin
    if Devise = V_PGI.DevisePivot then TR.PutValue('CET_CREDIT', Montant)
                                  else TR.PutValue('CET_CREDITDEV', Montant);
  end else begin
    if Devise = V_PGI.DevisePivot then TR.PutValue('CET_DEBIT', Abs(Montant))
                                  else TR.PutValue('CET_DEBITDEV', Abs(Montant));
  end;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_CETEBAC.RempliSoldeFinal(TR : TOB; StReleve : string; NumReleve : Integer);
{---------------------------------------------------------------------------------------}
var
  Montant : double;Devise : string ;
begin
  Inc(NumLigne);

  TR.PutValue('CET_NUMRELEVE', NumReleve);
  TR.PutValue('CET_ASUPPRIMER', '-');
  TR.PutValue('CET_TYPELIGNE', '07' );
  TR.PutValue('CET_NUMLIGNE', NumLigne);
  TR.PutValue('CET_ETABBQ', Copy(StReleve, 3, 5));
  TR.PutValue('CET_GUICHET', Copy(StReleve, 12, 5));
  TR.PutValue('CET_NUMEROCOMPTE', Copy(StReleve, 22, 11));
  TR.PutValue('CET_CODEAFB', '');
  TR.PutValue('CET_DATEOPERATION', ConvertDate(Copy(StReleve, 35, 6)));
  TR.PutValue('CET_CONTROLOK', '-');

  TR.PutValue('CET_EXONERE', 0);
  TR.PutValue('CET_IMO', '-');{24/04/07 : FQ TRESO 10441 : pour Oracle}
  TR.PutValue('CET_DATEVALEUR', ConvertDate(Copy(StReleve, 35, 6)));
  TR.PutValue('CET_LIBELLE', '');
  TR.PutValue('CET_LIBELLE1', '');
  TR.PutValue('CET_LIBELLE2', '');
  TR.PutValue('CET_LIBELLE3', '');
  TR.PutValue('CET_NUMEROPIECE', '');
  TR.PutValue('CET_REFPIECE', '');

  Devise := Copy(StReleve, 17, 3);
  TR.PutValue('CET_DEVISE', Devise);

  TR.SetDouble('CET_CREDIT'   , 0);
  TR.SetDouble('CET_DEBIT'    , 0);
  TR.SetDouble('CET_CREDITDEV', 0);
  TR.SetDouble('CET_DEBITDEV' , 0);

  {Ici, il s'agit du nouveau solde}
  FDecimal := StrToInt(Copy(StReleve, 20, 1));
  Montant := ConvertMontant(Copy(StReleve, 91, 14), FDecimal);
  if Montant > 0 then begin
    if Devise = V_PGI.DevisePivot then TR.PutValue('CET_CREDIT', Montant)
                                  else TR.PutValue('CET_CREDITDEV', Montant);
  end else begin
    if Devise = V_PGI.DevisePivot then TR.PutValue('CET_DEBIT', Abs(Montant))
                                  else TR.PutValue('CET_DEBITDEV', Abs(Montant));
  end;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_CETEBAC.RempliLigne(TR : TOB; StReleve : string; NumReleve : Integer);
{---------------------------------------------------------------------------------------}
var
  Montant : Double;
  Devise  : string;
begin
  Inc(NumLigne);

  TR.PutValue('CET_NUMRELEVE', NumReleve);
  TR.PutValue('CET_ASUPPRIMER', '-');
  TR.PutValue('CET_TYPELIGNE', '04' );
  TR.PutValue('CET_NUMLIGNE', NumLigne);
  TR.PutValue('CET_ETABBQ', Copy(StReleve, 3, 5));
  TR.PutValue('CET_GUICHET', Copy(StReleve, 12, 5));
  TR.PutValue('CET_NUMEROCOMPTE', Copy(StReleve, 22, 11));
  TR.PutValue('CET_CODEAFB', Copy(StReleve, 33, 2));
  TR.PutValue('CET_DATEOPERATION', ConvertDate(Copy(StReleve, 35, 6)));

  if Copy(StReleve, 89, 1)<> ' ' then TR.PutValue('CET_EXONERE', Copy(StReleve, 89, 1));
  TR.SetDateTime('CET_DATEVALEUR', ConvertDate(Copy(StReleve, 43, 6)));
  TR.SetString('CET_LIBELLE'     , Copy(StReleve, 49, 31));
  TR.SetString('CET_REFPIECE'    , Copy(StReleve, 105, 16));
  TR.SetString('CET_NUMEROPIECE' , Copy(StReleve, 82, 7));
  {24/04/07 : FQ TRESO 10441 : pour Oracle}
  if Copy(StReleve, 21, 1) = '' then
    TR.SetString('CET_IMO'       , '-')
  else
    TR.SetString('CET_IMO'       , Copy(StReleve, 21, 1));
  TR.SetString('CET_LIBELLE1'    , '');
  TR.SetString('CET_LIBELLE2'    , '');
  TR.SetString('CET_LIBELLE3'    , '');
  TR.SetString('CET_CONTROLOK'   , '-');

  Devise := Copy(StReleve, 17, 3);
  TR.SetString('CET_DEVISE', Devise);

  TR.SetDouble('CET_CREDIT'   , 0);
  TR.SetDouble('CET_DEBIT'    , 0);
  TR.SetDouble('CET_CREDITDEV', 0);
  TR.SetDouble('CET_DEBITDEV' , 0);

  Montant := ConvertMontant(Copy(StReleve, 91, 14), StrToInt(Copy(StReleve, 20, 1)));
  if Montant > 0 then begin
    if Devise = V_PGI.DevisePivot then TR.PutValue('CET_CREDIT', Montant)
                                  else TR.PutValue('CET_CREDITDEV', Montant);
  end else begin
    if Devise = V_PGI.DevisePivot then TR.PutValue('CET_DEBIT', - 1 * Montant)
                                  else TR.PutValue('CET_DEBITDEV', - 1 * Montant);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CETEBAC.RempliLibelleComplementaire(TR, TRLigne: TOB; StReleve: string);
{---------------------------------------------------------------------------------------}
var
  TCh : TOB;
begin
  Tch := TRLigne;

       if (Tch.GetString('CET_LIBELLE1') = '') then Tch.SetString('CET_LIBELLE1', Trim(Copy(StReleve, 41, 41)))
  else if (Tch.GetString('CET_LIBELLE2') = '') then Tch.SetString('CET_LIBELLE2', Trim(Copy(StReleve, 41, 41)))
  else if (Tch.GetString('CET_LIBELLE3') = '') then Tch.SetString('CET_LIBELLE3', Trim(Copy(StReleve, 41, 41)));
end;

{JP 21/09/04 : FQ TRESO 10120 : possibilité d'imprimer plusieurs relevés
{---------------------------------------------------------------------------------------}
procedure TOF_CETEBAC.ImprimerClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  T : TOB;
begin
  {Si aucune ligne sélectionnées}
  if (HGFICHIER.nbSelected = 0) and not HGFICHIER.AllSelected then begin
    PGIError(TraduireMemoire('Veuillez sélectionner un relevé.'), Ecran.Caption);
    Exit;
  end;

  T := TOB.Create('$$$', nil, -1);
  try
    {Constitution de la tob d'impression}
    for n := 0 to HGFICHIER.RowCount - 1 do
      if HGFICHIER.IsSelected(n)then
        RemplitTobImp(HGFICHIER.Cells[cColNomFichier, n], T);
    {Lancement de l'état}
    LanceEtatTob('E', 'ECT', 'REL', T, True, False, False, nil, '', TraduireMemoire('Intégration de relevés bancaires'), False);
  finally
    FreeAndNil(T);
  end;
end;

{JP 21/09/04 : FQ TRESO 10120 : Remplit la Tob d'impression à partir d'un fichier de relevé
{---------------------------------------------------------------------------------------}
procedure TOF_CETEBAC.RemplitTobImp(Fic : string; var T : TOB);
{---------------------------------------------------------------------------------------}
var
  D : TOB;
  F : TextFile;
  Q : TQuery;
  Ligne : string;
  Code  : string;
  Cpte  : string;
  Deci  : Integer;
  Montant : Double;
begin
  if (Fic = '') or not Assigned(T) then Exit;

  AssignFile(F, Fic);
  Reset(F);
  Cpte := '';

  try
    while not EOF(F) do begin
      Readln(F, Ligne);
      {Recherche du libellé du compte à partir du RIB}
      if Cpte = '' then begin
        Q := OpenSQL('SELECT BQ_LIBELLE FROM BANQUECP WHERE BQ_ETABBQ = "' + Copy(Ligne, 3, 5) + '" AND ' +
                     'BQ_GUICHET = "' + Copy(Ligne, 12, 5) + '" AND BQ_NUMEROCOMPTE = "' + Copy(Ligne, 22, 11) + '"', True,-1,'',true);
        if not Q.EOF then Cpte := Q.FindField('BQ_LIBELLE').AsString
                     else Cpte := '    ';
        Ferme(Q);
      end;

      Code := Copy(Ligne, 1, 2);
      if (Code = '01') or (Code = '04') or (Code = '07') then begin
        D := TOB.Create('$$$', T, -1);
        {Gestion du type de ligne, du libellé du compte et de son RIB}
        D.AddChampSupValeur('CODE', Code);
        D.AddChampSupValeur('COMPTE', Cpte);
        D.AddChampSupValeur('BRIB', Copy(Ligne, 3, 5) + ' ' + Copy(Ligne, 12, 5) + ' ' + Copy(Ligne, 22, 11));

        {Gestion des champs DEBIT et CREDIT}
        Deci := StrToInt(Copy(Ligne, 20, 1));
        Montant := ConvertMontant(Copy(Ligne, 91, 14), Deci);
        D.AddChampSup('DEBIT', False);
        D.AddChampSup('CREDIT', False);
             if Montant < 0 then D.SetString('DEBIT' , FloatToStrF(Montant * -1, ffFixed, 20, Deci))
        else if Montant > 0 then D.SetString('CREDIT', FloatToStrF(Montant, ffFixed, 20, Deci))
        else if Montant = 0 then D.SetString('CREDIT', '0.00');

        {Ajout de la date du relevé et de la devise}
        D.AddChampSupValeur('DATERELEV', DateToStr(ConvertDate(Copy(Ligne, 35, 6))));
        D.AddChampSupValeur('DEVISE', Copy(Ligne, 17, 3));

        {Gestion du libellé des lignes et de la date de valeur}
        D.AddChampSup('DATEVALEUR', False);
             if (Code = '01') then D.AddChampSupValeur('LIBELLE', 'Solde Initial ')
        else if (Code = '07') then D.AddChampSupValeur('LIBELLE', 'Solde Final ')
        else begin
          D.AddChampSupValeur('LIBELLE', Copy(Ligne, 49, 31));
          D.SetString('DATEVALEUR', DateToStr(ConvertDate(Copy(Ligne, 43, 6))));
        end;
        end;
      end ;
  finally
    CloseFile(F);
  end ;
end;

{---------------------------------------------------------------------------------------}
function ConvertMontant(MontantChar : string; Deci : Integer) : Double ;
{---------------------------------------------------------------------------------------}
var
  MontantFloat : Double;
  Lettre : Char;
  Divi   : Integer;
  n      : Integer;
begin
  Lettre := MontantChar[14] ;
  Divi := 1;
  MontantFloat :=  1;

  for n := 1 to Deci do Divi := Divi * 10;

  case Lettre of
    '{'      : Lettre := '0';
    '}'      : begin
                 MontantFloat := -1;
                 Lettre := '0';
               end;
    'A'..'I' : Dec(Lettre, 16);
    'J'..'R' : begin
                 MontantFloat := -1;
                 Dec(Lettre, 25);
               end;
  end;
  Result := MontantFloat * Valeur(Copy(MontantChar, 1, 13) + Lettre) / Divi;
end;

{---------------------------------------------------------------------------------------}
function ConvertDate(DateChar : string) : TDateTime ;
{---------------------------------------------------------------------------------------}
var
  a, m, j : Word;
begin
  a  := StrToInt(Copy(DateChar, 5, 2));
  if a > 90 then a := a + 1900
            else a := a + 2000;
  m := StrToInt(Copy(DateChar, 3, 2));
  j := StrToInt(Copy(DateChar, 1, 2));
  Result:= EncodeDate(a, m, j);
end;

{Contrôle la validité du relevé :
  1/ Présence d'enregistrement 01 et 07
  2/ Le solde du 01 + Montants des 04 = Solde du 07
  3/ Contrôle de la continuité du relevé avec les relevés intégrés
     a/ le fichier n'a pas déjà été intégré
     b/ il n'y a pas de "trou" entre le dernier relevé intégré et celui en cours
  4/ On s'assure qu'il n'y a pas de session de pointage postérieure à la date du relevé
{---------------------------------------------------------------------------------------}
function TOF_CETEBAC.ControleReleve(var TobRel : TOB): Boolean;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  M : Double;
  T : Double;
  F : TOB;
  O : TOB;
  S : string;
  R : TRecRib;
begin
  Result := False;
  {1/ Fichier incomplet : un fichier doit comporter au moins un 01 et un 07}
  if TobRel.Detail.Count < 2 then begin
    if TobRel.Detail.Count = 0 then
      with Tob.Create('****', TobRel, -1) do
        AddChampSupValeur('ERREUR', TraduireMemoire('Relevé vide'))
    else
      TobRel.Detail[0].AddChampSupValeur('ERREUR', TraduireMemoire('Relevé incomplet'));
  end

  else if (TobRel.FindFirst(['CET_TYPELIGNE'], ['01'], True) = nil) then
    TobRel.Detail[0].AddChampSupValeur('ERREUR', TraduireMemoire('Enregistrement 01 absent'), True)
  else if (TobRel.FindFirst(['CET_TYPELIGNE'], ['07'], True) = nil) then
    TobRel.Detail[0].AddChampSupValeur('ERREUR', TraduireMemoire('Enregistrement 07 absent'), True)

  {Autres cas}
  else begin
    {2/ Contrôle de la cohérence des soldes : théoriquement ne devrait pas arrivé}
    m := 0;
    for n := 0 to TobRel.Detail.Count - 1 do begin
      F := TobRel.Detail[n];
      if F.GetString('CET_TYPELIGNE') <> '07' then
        m := m + F.GetDouble('CET_CREDITDEV') + F.GetDouble('CET_CREDIT') - F.GetDouble('CET_DEBITDEV') - F.GetDouble('CET_DEBIT');
    end;
    F := TobRel.FindFirst(['CET_TYPELIGNE'], ['07'], True);
    {On compare le cumul des ligne 01 et 04 avec la ligne 07}
    T := F.GetDouble('CET_CREDITDEV') + F.GetDouble('CET_CREDIT') - F.GetDouble('CET_DEBITDEV') - F.GetDouble('CET_DEBIT');
    if Arrondi(m, FDecimal) <> Arrondi(T, FDecimal) then
      TobRel.Detail[0].AddChampSupValeur('ERREUR', TraduireMemoire('Les soldes du fichier ne sont pas cohérents'), True)

    {Depuis le bureau, les contrôles sont faits lors de la création d'une session de pointage.
     En effet, il est impossible de faire les contrôles car CETEBAC est une table DB000000}
    else if (ctxLanceur in V_PGI.PGIContexte) then
      Result := True

    {3/ Contrôle de l'intégration en cours avec EEXBQ}
    else begin
      O := TobRel.FindFirst(['CET_TYPELIGNE'], ['07'], True);
      F := TobRel.FindFirst(['CET_TYPELIGNE'], ['01'], True);
      m := F.GetDouble('CET_CREDITDEV') + F.GetDouble('CET_CREDIT') - F.GetDouble('CET_DEBITDEV') - F.GetDouble('CET_DEBIT');
      R.Etab    := F.GetString('CET_ETABBQ');
      R.Guichet := F.GetString('CET_GUICHET');
      R.Compte  := F.GetString('CET_NUMEROCOMPTE');

      S := HasCoherenceCEtebacEexbq(F.GetDateTime('CET_DATEOPERATION'), O.GetDateTime('CET_DATEOPERATION'), R, M);

      Result := (Trim(S) = '') or (Trim(S) = SANSSESSIONPTGE);

      if (Trim(S) <> '') and (Trim(S) <> SANSSESSIONPTGE) then
        TobRel.Detail[0].AddChampSupValeur('ERREUR', S, True)

      else if (Trim(S) = '') then begin
        {Si S = SANSSESSIONPTGE, cela signifie qu'aucune session de pointage antérieure n'a été trouvée.
         De ce fait le contrôles sur les soldes n'a pas été effectué : on ne met pas CET_CONTROLOK à X, de
         manière à ce que les contrôles soient relancés dans la uTomEEXBQ, lors de l'intégration de CETEBAC
         dans EEXBQLIG. En effet, on a pu ici intégrer plusieurs relevés du même compte qui ne seront intégrés
         dans EEXBQLIG que un par un ; il sera alors possible de vérifié la cohérence des soldes entre ces relevés}
        for n := 0 to TobRel.Detail.Count - 1 do
          TobRel.Detail[n].SetString('CET_CONTROLOK', 'X');
      end;
    end;
  end;
end;

{Sauvegarde du fichier dans le sous-répertoire Sauv
{---------------------------------------------------------------------------------------}
procedure TOF_CETEBAC.SauvegardeFichier(NomFichier : string; Fichier : TStringList; Del : Boolean);
{---------------------------------------------------------------------------------------}
var
  Nom : string;
  Rep : string;
begin
  try
    {Demande de sauvegarde : on copie le fichier dans Bak après avoir changer l'extension}
    if not Del then begin
      Rep := ExtractFilePath(NomFichier) + '\BAK';
      SysUtils.ForceDirectories(Rep);
      Nom := ExtractFileName(NomFichier);
      //ChangeFileExt(Nom, '.INT');
      Fichier.SaveToFile(Rep + '\' + Nom);
    end;

    {Suppression du fichier de son répertoire d'origine}
    SysUtils.DeleteFile(NomFichier);
  except
    on E : Exception do
      PGIError(TraduireMemoire('Une erreur est intervenur lors de la gestion des fichiers avec le message :') +
               #13#13 + E.Message, Ecran.Caption);
  end;
end;

initialization
  RegisterClasses([TOF_CETEBAC]);

end.

