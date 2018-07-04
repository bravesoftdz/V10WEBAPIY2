unit UTofMBOArreteStock;

interface

uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
  HCtrls, HEnt1, HMsgBox, UTOF, UTOB, Spin,
  {$IFDEF EAGLCLIENT}
  eQRS1, utileAGL, MaineAGL, eFiche,
  {$ELSE}
  QRS1, db, dbTables, EdtEtat, Fiche, Fe_Main,
  {$ENDIF}
  AglInit, EntGC, HQry, HStatus, MajTable, UtilGc, graphics, M3FP, ParamSoc,
  HDimension, UDimArticle, UtilDimArticle, UtilArticle, Vierge, UtilPgi, HPanel,
  InvUtil, FactComm, FactUtil, voirtob;

type
  TOF_MBOArreteStock = class(TOF)
    DimensionsArticle: TODimArticle;
    procedure OnArgument(ArgumentS: string); override;
    procedure OnLoad; override;
    procedure OnUpdate; override;
    procedure OnClose; override;
  private
    StatutArt, MasqueDim: string;
    Article, CodeArticle: string;
    natureDoc: string;
    TobArt, TobArreteStock: Tob;
    CanClose: boolean;
    nbAff: integer;
    ChampAffiche: array[1..MaxDimChamp] of string; // param USERPREFDIM
    LibelleAffiche: array[1..MaxDimChamp] of string; // param USERPREFDIM
    //function GetPlusMoins(TypeQte,QtePlus,QteMoins : string) : integer ;
    procedure InitDimensionsArticle;
    procedure LoadParamUser;
    procedure AfficheParamUser(AvecDonnees: boolean);
    procedure AffEcran(ExisteStock: boolean);
  public
    procedure ParamTailleUnique;
  end;
function RecalculStockDate(StatutArt, Article, CodeArticle, DateStock: string; Depot: string = '???'; Depart: integer = 1; MvtJDinclus: boolean = False;
  MvtJFinclus: boolean = False; EcartINVinclus: boolean = False; TobRetour: tob = nil): boolean;

const
  TexteMessage: array[1..1] of string = (
    {1}'Pas de stock pour cet article dans cet établissement'
    );

implementation

const
  // Paramètre Départ : recalcul depuis clôture ou depuis stock actuel
  STOCK_CLOTURE = 1;
  STOCK_ACTUEL = 2;

procedure TOF_MBOArreteStock.OnArgument(Arguments: string);
var Critere, ChampMul, ValMul: string;
  x: integer;
begin
  inherited;

  CanClose := False;
  Critere := UpperCase(Trim(ReadTokenSt(Arguments)));
  while Critere <> '' do
  begin
    x := pos('=', Critere);
    if x <> 0 then
    begin
      ChampMul := copy(Critere, 1, x - 1);
      ValMul := copy(Critere, x + 1, length(Critere));
      if ChampMul = 'ARTICLE' then Article := ValMul;
    end;
    Critere := UpperCase(Trim(ReadTokenSt(Arguments)));
  end;

  DimensionsArticle := nil;
  // Chargement de l'article courant en tob.
  TobArt := TOB.Create('ARTICLE', nil, -1);
  TobArt.SelectDB('"' + Article + '"', nil);
  StatutArt := TobArt.GetValue('GA_STATUTART');
  CodeArticle := TobArt.GetValue('GA_CODEARTICLE');
  MasqueDim := TobArt.GetValue('GA_DIMMASQUE');

  TobArt.PutEcran(Ecran);
  SetControlText('DATESTOCK', DateToStr(V_PGI.DateEntree));
  SetControlText('_ETABLISS', GetParamSoc('SO_GCDEPOTDEFAUT'));
  natureDoc := NAT_ARRETESTO; // Fiche Arrêté de stock :  MBO ARRETESTOCK
  SetControlEnabled('BPARAM', False);
  //SetControlVisible('BPARAM',NaturePieceGeree(natureDoc)) ;
  SetControlVisible('BPARAM', False); // Paramétrage non opérationnel

  AffEcran(True);
  if StatutArt <> 'GEN' then
  begin
    // Mise en place des champs paramétrés dans UserPrefDim
    LoadParamUser;
    AfficheParamUser(False);
  end;
end;

procedure TOF_MBOArreteStock.OnLoad;
begin
  inherited;
end;

procedure TOF_MBOArreteStock.OnUpdate;
var iChamp: integer;
  DateDebIncluse, DateFinIncluse, CalculOk: boolean;
  DateStock, Depot: string;
begin
  inherited;
  // Ré-initialisation : cas de +sieurs lancements successifs.
  //if DimensionsArticle<>nil then DimensionsArticle.free;
  if DimensionsArticle <> nil then
  begin
    DimensionsArticle.Destroy;
    DimensionsArticle := nil;
    TobArreteStock := nil;
  end;

  DateDebIncluse := (TCheckBox(GetControl('CBINCLUS')).State = cbChecked);
  // DCA - FQ MODE 10717 - Prise en compte ou non des mvts d'aujourd'hui
  DateFinIncluse := (TCheckBox(GetControl('CBAUJOURDHUI')).State = cbChecked);
  DateStock := GetControlText('DATESTOCK');
  Depot := GetControlText('_ETABLISS');
  TobArreteStock := TOB.Create('DISPO', nil, -1);

  CalculOk := RecalculStockDate(StatutArt, Article, CodeArticle, DateStock, Depot, 1, DateDebIncluse, DateFinIncluse, False, TobArreteStock);

  if not CalculOk then
  begin
    AffEcran(False);
    // Message article sans stock
    //PGIBox(TexteMessage[1],Ecran.Caption)
  end else
  begin
    if GetControlVisible('TSANSSTOCK') then AffEcran(True);
    if StatutArt = 'GEN' then
    begin
      TheTob := TobArreteStock;
      InitDimensionsArticle;
      TheTob := nil;
    end
    else
      for iChamp := 1 to nbAff do SetControlText('CHAMP' + IntToStr(iChamp), TobArreteStock.Detail[0].GetValue(ChampAffiche[iChamp]));
  end;

  // Paramétrage disponible dès que la grille / les zones sont chargées
  //SetControlEnabled('BPARAM',True) ;

end;

procedure TOF_MBOArreteStock.OnClose;
begin
  inherited;
  if CanClose then
  begin
    LastError := 0;
    if DimensionsArticle <> nil then
    begin
      DimensionsArticle.Destroy;
      DimensionsArticle := nil
    end;
    if TobArt <> nil then
    begin
      TobArt.Free;
      TobArt := nil
    end;
    if (StatutArt = 'UNI') and (TobArreteStock <> nil)
      then
      begin
      TobArreteStock.Free;
      TobArreteStock := nil
    end;
  end
  else LastError := -1;
end;

procedure TOF_MBOArreteStock.InitDimensionsArticle;
var iItem: integer;
begin
  DimensionsArticle := TODimArticle.Create(THDimension(GetControl('THDIM')), GetControlText('GA_CODEARTICLE'),
    MasqueDim, '', 'GCDIMCHAMP', natureDoc, '', GetControlText('_ETABLISS'), '', False);
  for iItem := 0 to 3 do DimensionsArticle.Dim.PopUp.Items[iItem].Visible := False;
  DimensionsArticle.Dim.Action := TFFiche(Ecran).FTypeAction;

  AfficheUserPref(DimensionsArticle, natureDoc, '');
end;

procedure TOF_MBOArreteStock.LoadParamUser;
var position, iChamp: integer;
  stSql, stChamp: string;
  QQ: TQuery;
begin
  stSql := 'select GUD_POSITION,GUD_CHAMP,GUD_LIBELLE,GUD_DETAIL,GUD_DEFAUT,' +
    ' GUD_TYPEMASQUE,GUD_ETABLISSEMENT' +
    ' from USERPREFDIM' +
    ' where GUD_NATUREPIECE="' + NAT_ARRETESTO + '" and GUD_UTILISATEUR="' + V_PGI.User + '"' +
    ' and GUD_POSITION>0 and GUD_POSITION<11' +
    ' order by GUD_POSITION';
  QQ := OpenSQL(stSql, True);
  nbAff := 0;
  if not QQ.Eof then while not QQ.EOF do
    begin
      position := QQ.FindField('GUD_POSITION').AsInteger;
      inc(nbAff);
      if (position > 0) and (position < 11) then
      begin
        ChampAffiche[nbAff] := QQ.FindField('GUD_CHAMP').AsString;
        LibelleAffiche[nbAff] := QQ.FindField('GUD_LIBELLE').AsString;
      end;
      QQ.Next;
    end;
  Ferme(QQ);

  // Contrôle présence des champs obligatoires
  QQ := OpenSQL('select CO_ABREGE,CO_LIBELLE from COMMUN where CO_TYPE="GDO" and CO_LIBRE like "%' + NAT_ARRETESTO + '%" order by CO_CODE', True);
  while not QQ.EOF do
  begin
    if QQ.Findfield('CO_ABREGE').AsString <> '' then
    begin
      iChamp := 1;
      stChamp := QQ.Findfield('CO_ABREGE').AsString;
      while (stChamp <> ChampAffiche[iChamp]) and (iChamp < MaxDimChamp) do inc(iChamp);
      if (stChamp <> ChampAffiche[iChamp]) and (nbAff < MaxDimChamp) then
      begin
        // Ajout des champs obligatoires en fin de liste
        inc(nbAff);
        ChampAffiche[nbAff] := stChamp;
        LibelleAffiche[nbAff] := QQ.Findfield('CO_LIBELLE').AsString;
      end;
    end;
    QQ.next;
  end;
  ferme(QQ);

end;

procedure TOF_MBOArreteStock.AfficheParamUser(AvecDonnees: boolean);
var iAff: integer;
begin
  for iAff := 1 to MaxDimChamp do
  begin
    SetControlVisible('T_CHAMP' + IntToStr(iAff), iAff <= nbAff);
    SetControlVisible('CHAMP' + IntToStr(iAff), iAff <= nbAff);
    SetControlText('T_CHAMP' + IntToStr(iAff), LibelleAffiche[iAff]);
    if AvecDonnees and TobArreteStock.Detail[0].FieldExists(ChampAffiche[iAff])
      then SetControlText('CHAMP' + IntToStr(iAff), TobArreteStock.Detail[0].GetValue(ChampAffiche[iAff]));
  end;
end;

// NaturePiece affecte-t-elle la quantité TypeQte ?
{
function TOF_MBOArreteStock.GetPlusMoins(TypeQte,QtePlus,QteMoins : string) : integer ;
begin
if Pos(TypeQte,QtePlus)>0 then result:=1
else if Pos(TypeQte,QteMoins)>0 then result:=-1
else result:=0 ;
end;
}

procedure TOF_MBOArreteStock.AffEcran(ExisteStock: boolean);
begin
  if ExisteStock then
  begin
    SetControlProperty('TSANSSTOCK', 'Top', 50);
    SetControlVisible('TSANSSTOCK', False);
    if StatutArt <> 'GEN' then
    begin
      SetControlVisible('P_DIMENSION', False);
      SetControlVisible('P_TAILLEUNIQUE', True);
      Ecran.Height := 285;
      THPanel(GetControl('P_TAILLEUNIQUE')).Align := alBottom;
    end else
    begin
      SetControlVisible('P_TAILLEUNIQUE', False);
      SetControlVisible('P_DIMENSION', True);
      Ecran.Height := 400;
      SetControlProperty('P_DIMENSION', 'Height', 255);
      THPanel(GetControl('P_DIMENSION')).Align := alBottom;
    end;
  end else
  begin
    SetControlVisible('P_TAILLEUNIQUE', False);
    SetControlVisible('P_DIMENSION', False);
    SetControlVisible('TSANSSTOCK', True);
    SetControlProperty('TSANSSTOCK', 'Top', 130);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 21/01/2002
Modifié le ... : 03/03/2003
Description .. : Recalcul du stock à une date donnée
Suite ........ : Paramètres :
Suite ........ : - StatutArt : statut de l'article GEN,DIM,UNI.
Suite ........ : - Article : article en taille unique, dimensionné ou générique.
Suite ........ : - DateStock : date à laquelle on souhaite connaître le
Suite ........ : stock.
Suite ........ : - Dépôt : (facultatif) dépôt du stock interrogé.
Suite ........ : - Départ : (facultatif) précise si le recalcul commence à la
Suite ........ : première clôture postérieure
Suite ........ : ou si l'on repart du stock actuel. Par défaut, recalcul depuis
Suite ........ : clôture.
Suite ........ : - MvtJDinclus : (facultatif) le stock recalculé inclu-t-il
Suite ........ : les mouvements du jour de l'arrêté ?
Suite ........ : - MvtJFinclus : (facultatif) le stock recalculé inclu-t-il
Suite ........ : les mouvements du jour de la clôture ?
Suite ........ : - EcartINVinclus : (facultatif) Inclure ou non les écarts
Suite ........ : d'inventaire.
Suite ........ : - TobRetour : (facultatif) si une tob est renseignée, tous les
Suite ........ : compteurs (quantités, pmap et prmp)
Suite ........ : seront recalculés et renseignés dans cette tob. Par défaut,
Suite ........ : pas de Tob.
Suite ........ : Retour :
Suite ........ : - True si calcul ok / False si article absent de la table
Suite ........ : DISPO,
Suite ........ : - Tous les compteurs (future version : pmap et pmrp) de la
Suite ........ : fiche
Suite ........ : stock via le paramètre TobRetour s'il est défini.
Mots clefs ... : ARTICLE;STOCK;RECALCUL;ARRETE
*****************************************************************}

function RecalculStockDate(StatutArt, Article, CodeArticle, DateStock: string; Depot: string = '???'; Depart: integer = 1; MvtJDinclus: boolean = False;
  MvtJFinclus: boolean = False; EcartINVinclus: boolean = False; TobRetour: tob = nil): boolean;
var QQ, QDispo, QLigne, QPrec: TQuery;
  LigneATraiter: boolean; //,LigneVivante : boolean ;
  NbLig, Sens, SensPrec, iCpt: integer;
  QtePrec: double;
  stWhereArt, stDateMax, stSql, stNature, stMvtJDinclus, stMvtJFinclus, stEcartINVinclus: string;
  stPiecePrecedente, stSqlPrec: string; //,stPiece : string ;
  CleDocPrec: R_CleDoc;
  TobDispo, TobD, TobPieceNonTraitee, TobPNT: TOB;

const TabCptr: array[1..12, 1..2] of string = (
    {1}('PHY', 'GQ_PHYSIQUE')
    {2}, ('ESE', 'GQ_ENTREESORTIES')
    {3}, ('LB1', 'GQ_QTE1')
    {4}, ('LB2', 'GQ_QTE2')
    {5}, ('LB3', 'GQ_QTE3')
    {6}, ('LC', 'GQ_LIVRECLIENT')
    {7}, ('LF', 'GQ_LIVREFOU')
    {8}, ('PRE', 'GQ_PREPACLI')
    {9}, ('RC', 'GQ_RESERVECLI')
    {10}, ('RF', 'GQ_RESERVEFOU')
    {11}, ('TRA', 'GQ_TRANSFERT')
    {12}, ('VFO', 'GQ_VENTEFFO')
    );

  MAXCPTR = 12;

  procedure MajCompteur(stCpt, NomChamp: string);
  begin
    Sens := GetPlusMoinsCpt(stNature, stCpt);
    if (Sens <> 0) then
    begin
      QtePrec := 0.0;
      if stPiecePrecedente <> '' then
      begin
        SensPrec := GetPlusMoinsCpt(CleDocPrec.NaturePiece, stCpt);
        if Sens = SensPrec then // Début des ennuis
        begin
          stSqlPrec := 'select GL_QTEFACT,GL_QTERESTE from LIGNE ' +
            'where GL_NATUREPIECEG="' + CleDocPrec.NaturePiece +
            '" and GL_SOUCHE="' + CleDocPrec.Souche +
            '" and GL_NUMERO=' + IntToStr(CleDocPrec.NumeroPiece) +
            ' and GL_INDICEG=' + IntToStr(CleDocPrec.Indice) +
            ' and GL_NUMLIGNE=' + IntToStr(CleDocPrec.NumLigne);
          QPrec := OpenSql(stSqlPrec, True);
          if not QPrec.Eof then
          begin
            QtePrec := QPrec.FindField('GL_QTEFACT').AsFloat -
              QPrec.FindField('GL_QTERESTE').AsFloat;
            if CleDocPrec.NaturePiece = stNature then QtePrec := -QtePrec;
          end;
          Ferme(QPrec);
        end;
      end;
      TobD.PutValue(NomChamp,
        TobD.GetValue(NomChamp) - ((QLigne.FindField('GL_QTEFACT').AsFloat - QtePrec)) * Sens);
    end;
  end; // Procedure MajCompteur(stCpt,NomChamp : string) ;

begin
  Result := False;

  if CodeArticle = '???' then ; // Lire fiche article
  if Depot = '???' then Depot := GetParamSoc('SO_GCDEPOTDEFAUT');
  if StatutArt <> 'GEN' then stWhereArt := '_ARTICLE="' + Article + '"'
  else stWhereArt := '_ARTICLE in (select GA_ARTICLE from ARTICLE where GA_STATUTART="DIM" ' +
    'and GA_CODEARTICLE="' + CodeArticle + '")';

  // Recherche première clôture suivant la date d'arrêté de stock
  QDispo := nil;
  if Depart = STOCK_CLOTURE then
  begin
    stSql := 'select GQ_ARTICLE,GQ_DATECLOTURE,GQ_PHYSIQUE,GQ_RESERVECLI,GQ_RESERVEFOU,' +
      'GQ_PREPACLI,GQ_LIVRECLIENT,GQ_LIVREFOU,GQ_TRANSFERT,GQ_CUMULSORTIES,' +
      'GQ_CUMULENTREES,GQ_VENTEFFO,GQ_ECARTINV,GQ_PMAP,GQ_PMRP, ' +
      'GA_ARTICLE,GA_CODEARTICLE,GA_DIMMASQUE,GA_PCB, ' +
      'GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5, ' +
      'GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,GA_GRILLEDIM4,GA_GRILLEDIM5 ' +
      'from DISPO ' +
      'left join ARTICLE on GA_ARTICLE=GQ_ARTICLE ' +
      'where GQ' + stWhereArt + ' and GQ_DEPOT="' + Depot + '" and GQ_CLOTURE="X" ' +
      'and GQ_DATECLOTURE=(select min(GQ_DATECLOTURE) from DISPO where GQ' + stWhereArt +
      ' and GQ_DEPOT="' + Depot + '" and GQ_CLOTURE="X"' +
      ' and GQ_DATECLOTURE>="' + USDateTime(StrToDate(DateStock)) + '")';
    QDispo := OpenSQL(stSql, True);
    // Si pas de clôture trouvée, on repart du stock actuel.
    if QDispo.EOF then
    begin
      Depart := STOCK_ACTUEL;
      Ferme(QDispo)
    end
    else stDateMax := QDispo.findField('GQ_DATECLOTURE').AsString;
  end;
  if Depart = STOCK_ACTUEL then
  begin
    stSql := 'select GQ_ARTICLE,GQ_DATECLOTURE,GQ_PHYSIQUE,GQ_RESERVECLI,GQ_RESERVEFOU,' +
      'GQ_PREPACLI,GQ_LIVRECLIENT,GQ_LIVREFOU,GQ_TRANSFERT,GQ_QTE1,GQ_QTE2,' +
      'GQ_QTE3,GQ_VENTEFFO,GQ_ENTREESORTIES,GQ_ECARTINV,GQ_PMAP,GQ_PMRP ' +
      'GA_ARTICLE,GA_CODEARTICLE,GA_DIMMASQUE,GA_PCB, ' +
      'GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5, ' +
      'GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,GA_GRILLEDIM4,GA_GRILLEDIM5 ' +
      'from DISPO ' +
      'left join ARTICLE on GA_ARTICLE=GQ_ARTICLE ' +
      'where GQ' + stWhereArt + ' and GQ_DEPOT="' + Depot + '" and GQ_CLOTURE="-"';
    QDispo := OpenSQL(stSql, True);
    stDateMax := DateToStr(V_PGI.DateEntree);
  end;
  if QDispo = nil then exit
  else
    if QDispo.EOF then
  begin
    Ferme(QDispo);
    exit
  end;
  TobDispo := TOB.Create('DISPO', nil, -1);
  TobDispo.LoadDetailDB('DISPO', '', '', QDispo, False);
  Ferme(QDispo);

  // Mvts du jour d'arrêté de stock inclus dans le recalcul
  if MvtJDinclus then stMvtJDinclus := '>=' else stMvtJDinclus := '>';

  // Mvts du jour de clôture de stock inclus dans le recalcul
  if MvtJFinclus then stMvtJFinclus := '<=' else stMvtJFinclus := '<';

  // Ecarts d'inventaire inclus => Si True, pièces de nature INV ignorées !
  if EcartINVinclus then stEcartINVinclus := '' else stEcartINVinclus := ' and GP_NATUREPIECEG<>"INV"';

  if StatutArt <> 'GEN' then stWhereArt := '_ARTICLE="' + Article + '"'
  else stWhereArt := '_ARTICLE in (select GA_ARTICLE from ARTICLE where GA_STATUTART="DIM" ' +
    'and GA_CODEARTICLE=(select GA_CODEARTICLE from ARTICLE where GA_ARTICLE="' + Article + '"))';

  // Recherche nombre de lignes traitées -> pour la fenêtre de patience
  stSql := 'select count(*) as NBLIG ' +
    'from PIECE ' +
    'left join LIGNE on GL_NATUREPIECEG=GP_NATUREPIECEG and GL_SOUCHE=GP_SOUCHE and GL_NUMERO=GP_NUMERO and GL_INDICEG=GP_INDICEG ' +
    'where GP_NATUREPIECEG in (select GPP_NATUREPIECEG from PARPIECE where (GPP_QTEMOINS like "%PHY%" or GPP_QTEPLUS like "%PHY%"))' +
    stEcartINVinclus +
    ' and GP_DATEPIECE' + stMvtJDinclus + '"' + USDateTime(StrToDate(DateStock)) + '"' +
    ' and GP_DATEPIECE' + stMvtJFinclus + '"' + USDateTime(StrToDate(stDateMax)) + '"' +
    ' and GL' + stWhereArt +
    ' and GL_DEPOT="' + Depot + '"' +
    ' and GL_TENUESTOCK="X" and GL_QUALIFMVT<>"ANN" ';
  QQ := OpenSQL(stSql, True);
  if not QQ.Eof then NbLig := QQ.FindField('NBLIG').AsInteger else NbLig := 0;
  Ferme(QQ);

  stSql := 'select GP_NATUREPIECEG,GP_DATEPIECE,GP_NUMERO,GP_SOUCHE,GP_INDICEG,' +
    'GL_NUMLIGNE,GL_PIECEPRECEDENTE,GL_ARTICLE,GL_QTEFACT,GL_VIVANTE ' +
    //'GL_DPA,GL_DPR,GL_QUALIFQTESTO,GL_QUALIFQTEACH,GL_QUALIFQTEVTE ' +
  'from PIECE ' +
    'left join LIGNE on GL_NATUREPIECEG=GP_NATUREPIECEG and GL_SOUCHE=GP_SOUCHE and GL_NUMERO=GP_NUMERO and GL_INDICEG=GP_INDICEG ' +
    'where GP_NATUREPIECEG in (select GPP_NATUREPIECEG from PARPIECE where (GPP_QTEMOINS like "%PHY%" or GPP_QTEPLUS like "%PHY%"))' +
    stEcartINVinclus +
    ' and GP_DATEPIECE' + stMvtJDinclus + '"' + USDateTime(StrToDate(DateStock)) + '"' +
    ' and GP_DATEPIECE' + stMvtJFinclus + '"' + USDateTime(StrToDate(stDateMax)) + '"' +
    ' and GL' + stWhereArt +
    ' and GL_DEPOT="' + Depot + '"' +
    ' and GL_TENUESTOCK="X" and GL_QUALIFMVT<>"ANN" ' +
    'order by GP_DATEPIECE desc';
  //       ' order by GL_DATEPIECE desc, GP_DATECREATION desc, GP_HEURECREATION desc' ;
  QLigne := OpenSQL(stSql, True);

  TobPieceNonTraitee := Tob.Create('Pieces Non Traitees', nil, -1);

  InitMove(NbLig, '');
  QLigne.First;
  while not QLigne.Eof do
  begin
    // Recherche TobDispo de l'article
    TobD := TobDispo.FindFirst(['GQ_ARTICLE'], [QLigne.FindField('GL_ARTICLE').AsString], False);
    //LigneVivante:=QLigne.FindField('GL_VIVANTE').AsString='X' ;
    stNature := QLigne.FindField('GP_NATUREPIECEG').AsString;

    // Traitements des seules lignes VIVANTES pour CdeFou,LivreFou,CdeCli,PrepaCli,LivreCli ;
    //             de toutes les lignes pour les autres natures de pièce.
    LigneATraiter := (TobD <> nil);
    //if (not LigneVivante) and (pos(stNature,'CF;BLF;CC;PRE;BLC')>0) then LigneATraiter:=False ;
    stPiecePrecedente := QLigne.FindField('GL_PIECEPRECEDENTE').AsString;
    if LigneATraiter and (stPiecePrecedente <> '') then
    begin
      {        stPiece:=QLigne.FindField('GP_NATUREPIECEG').AsString + ';' +
                       QLigne.FindField('GP_SOUCHE').AsString + ';' +
                       QLigne.FindField('GP_NUMERO').AsString + ';' +
                       QLigne.FindField('GP_INDICEG').AsString + ';' +
                       QLigne.FindField('GL_NUMLIGNE').AsString + ';' ;
      }
              // Ne pas traiter les pièces d'indice supérieur = pièce précédante déjà traitée
      TobPNT := TobPieceNonTraitee.FindFirst(['PIECE'], [stPiecePrecedente], False);
      if TobPNT <> nil then LigneATraiter := False // Pièce précédente déjà traitée !!
      else
      begin
        TobPNT := TOB.Create('Pieces Non Traitees', TobPieceNonTraitee, -1);
        TobPNT.AddChampSup('PIECE', False);
        TobPNT.PutValue('PIECE', stPiecePrecedente);
      end;
    end;

    if LigneATraiter then
    begin
      //stPiecePrecedente:=QLigne.FindField('GL_PIECEPRECEDENTE').AsString ;
      if stPiecePrecedente <> '' then
        DecodeRefPiece(stPiecePrecedente, CleDocPrec);
      {
                  BEGIN
                  TobPPrec:=TobPiecePrec.FindFirst(['PIECEPREC'],[stPiecePrecedente],False) ;
                  if TobPPrec<>nil then stPiecePrecedente:=''  // Pièce précédente déjà traitée !!
                  else BEGIN
                       TobPPrec:=TOB.Create('Piece Prec',TobPiecePrec,-1) ;
                       TobPPrec.AddChampSup('PIECEPREC',False) ;
                       TobPPrec.PutValue('PIECEPREC',stPiecePrecedente) ;
                       DecodeRefPiece(stPiecePrecedente,CleDocPrec) ;
                       END ;
                  END ;
      }

      {
              // Maj GQ_PMAP et GQ_PMRP si le mvt a incrémenté le stock : A FAIRE ...
              --> GPP_MAJPRIXVALO
      }
      for iCpt := 1 to MAXCPTR do MajCompteur(TabCptr[iCpt, 1], TabCptr[iCpt, 2]);

    end;

    MoveCur(False);
    QLigne.next;
  end;
  FiniMove;
  Ferme(QLigne);
  TobPieceNonTraitee.Free;

  if TobRetour <> nil then TobRetour.Dupliquer(TobDispo, True, True, True);
  TobDispo.Free;
  Result := True;
end;

procedure TOF_MBOArreteStock.ParamTailleUnique;
var stInit, stRetour, Critere, ChampMul, ValMul: string;
  iChamp, iFor, iAff, x: integer;
begin
  stInit := NAT_ARRETESTO + ';;;;;;';
  for iChamp := 1 to MaxDimChamp do stInit := stInit + ChampAffiche[iChamp] + ';';
  stRetour := AglLanceFiche('MBO', 'ARTDIM_PARAM', '', '', stInit);
  nbAff := 0;
  iAff := 0;
  while stRetour <> '' do
  begin
    Critere := Trim(ReadTokenSt(stRetour));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then ChampMul := copy(Critere, 1, x - 1);
      ValMul := copy(Critere, x + 1, length(Critere));
      if ChampMul = 'VALEUR' then
      begin
        if nbAff = 0 then for iFor := 1 to MaxDimChamp do ChampAffiche[iFor] := '';
        if ValMul <> '' then
        begin
          inc(nbAff);
          ChampAffiche[nbAff] := ValMul;
        end;
      end
      else
        if (ChampMul = 'TITRE') and (ValMul <> '(Aucun)')
        then
        begin
        inc(iAff);
        LibelleAffiche[iAff] := ValMul
      end;
    end;
  end;
  AfficheParamUser(True);
end;

procedure AGLTOF_MBOArreteStock_OnClose(Parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then TOTOF := TFVierge(F).LaTOF else exit;
  if (TOTOF is TOF_MBOArreteStock) then TOF_MBOArreteStock(TOTOF).CanClose := StringToCheck(Parms[1]);
end;

procedure AGLTOF_MBOArreteStock_ParamGrille(Parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
  //ItemDim : THDimensionItem ;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then TOTOF := TFVierge(F).LaTOF else exit;
  if (TOTOF is TOF_MBOArreteStock) then
  begin
    with TOF_MBOArreteStock(TOTOF) do
    begin
      if StatutArt = 'GEN' then
      begin
        //itemDim:=DimensionsArticle.Dim.CurrentItem ;
        ParamGrille(DimensionsArticle, NAT_ARRETESTO, '');
        //if itemDim<>nil then ItemDimSetFocus(itemDim,'') ;
      end
      else ParamTailleUnique;
    end
  end;
end;

procedure InitTofMBOArreteStock();
begin
  RegisterAglProc('TOF_MBOArreteStock_OnClose', True, 1, AGLTOF_MBOArreteStock_OnClose);
  RegisterAglProc('TOF_MBOArreteStock_ParamGrille', True, 1, AGLTOF_MBOArreteStock_ParamGrille);
end;

initialization
  registerclasses([TOF_MBOArreteStock]);
  InitTofMBOArreteStock();

end.
