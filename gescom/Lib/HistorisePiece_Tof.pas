{***********UNITE*************************************************
Auteur  ...... : Dominique BROSSET
Créé le ...... : 05/12/2001
Modifié le ... : 09/09/2002
Description .. : Historisation des pièces
Mots clefs ... : PIECE;VIVANTE;HISTORISE
*****************************************************************}
unit HistorisePiece_Tof;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, HDimension,UTOM,AGLInit,
      Utob,HDB,Messages,HStatus,uTofAfBaseCodeAffaire, StockUtil,
{$IFDEF EAGLCLIENT}
      UtileAGL,eMul,MaineAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fiche, mul, db,DBGrids, EdtDoc, EdtEtat, Fe_Main,
{$endIF}
      M3VM, M3FP, Hqry, EntGC, FactCpta, FactComm, FactUtil, Ent1,
      UtilPGI, FactNomen;

function GCLanceFiche_HistorisePiece(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
     TOF_HistorisePiece = Class (TOF_AFBASECODEAFFAIRE)
        procedure OnArgument (stArgument : string); override;
        procedure OnUpdate; override;
     public
        TobPiece, TobLigne, TobNomenclature, TobNomenAPlat : TOB;
        bHistorique : boolean;
        stVentAchat : string;
        procedure AffecteDispo;
        procedure ChargeLesSousNomen (TobGroupeNomen, TobNomen : TOB;
                                      LaLig, MaxNiv, iDep : integer);
        procedure ChargeNomenclature (iIndPiece : integer);
        function ConstruitWhere (stPrefixe : string; TobPiece : TOB) : string;
        procedure Historisation;
        procedure LanceHistorisation ;
        procedure ModifieTableDispo (TobL, TobN : TOB;
                                     stArticle, stChamp : string;
                                     stDepot : string;
                                     iCoef : integer; dQte : double);
        procedure ModifieTableDispoContreM (TobL: Tob; stChamp : string);
        procedure ModifieTableDispoSerie (TobSerie : TOB;
                                          stDepot,
                                          stChamp : string);
    end ;

procedure AGLHistorisePiece (parms:array of variant; nb: integer);

const
// libellés des messages
TexteMessage: array[1..2] of string 	= (
          {1}  'Aucun élément sélectionné'
          {2} ,'Attention ! l''opération est irréversible' + #13 + 'Continuer?'
              );

implementation

function GCLanceFiche_HistorisePiece(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
Result:='';
if Nat='' then exit;
if Cod='' then exit;
Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

{==============================================================================================}
{================================== Procédure de la TOF =======================================}
{==============================================================================================}
procedure TOF_HistorisePiece.OnArgument (stArgument : string) ;
var stPlus : string;
    CC     : THValComboBox ;
begin
inherited ;
stVentAchat := stArgument; // appel avec Argument égal VEN ou ACH
stplus := ' AND GPP_VENTEACHAT="' + stArgument + '"'; // document de vente ou d'achat
stPlus := stPlus + ' AND GPP_ACTIONFINI="TRA"'; // que les documents qui se transforment
stPlus := stPlus + // les documents qui n'affectent pas le stock phy ou qui s'historise
        ' AND ((GPP_QTEPLUS not like "%PHY%" AND GPP_QTEMOINS not like "%PHY%") OR ' +
            'GPP_HISTORIQUE="X")'; // pour ne pas deleter des doc qui affectent le stock phy
//*** Blocages des natures de pièces autorisées en affaires
if (ctxAffaire in V_PGI.PGIContexte) and (stArgument='ACH') then
   begin
   if ctxScot in V_PGI.PGIContexte then
      begin
      SetControlText('GP_NATUREPIECEG','CF');
      Setcontrolenabled ('GP_NATUREPIECEG',False);
      end else
      begin
      stplus:=stPlus + ' AND ((GPP_NATUREPIECEG="CF") or (GPP_NATUREPIECEG="BLF") or (GPP_NATUREPIECEG="FF"))';
      end;
   end;
// **** Fin affaire ***
if not(ctxScot in V_PGI.PGIContexte) then setControlproperty ('GP_NATUREPIECEG','Plus',stplus);

if stArgument = 'VEN' then
    begin
    SetControlText ('GP_NATUREPIECEG', 'DE');
    SetControlCaption ('TGP_TIERS', 'Client');
    end else
    begin
    if Not (ctxScot in V_PGI.PGIcontexte) then SetControlText ('GP_NATUREPIECEG', 'CF');
    setControlproperty ('GP_TIERS','DataType','GCTIERSFOURN');
    setControlproperty ('GP_TIERS_','DataType','GCTIERSFOURN');
    SetControlText ('TGP_TIERS', 'Fournisseur');
    SetControlVisible ('TGP_TIERSPAYEUR', false);
    SetControlVisible ('GP_TIERSPAYEUR', false);
    end;
CC:=THValComboBox(GetControl('GP_DOMAINE'))       ; if CC<>Nil then PositionneDomaineUser(CC) ;
CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ; if CC<>Nil then PositionneEtabUser(CC) ;
end;

procedure TOF_HistorisePiece.OnUpdate;
begin
if GetInfoParpiece (GetControlText ('GP_NATUREPIECEG'), 'GPP_VENTEACHAT') <> stVentAchat then
    begin
    if stVentAchat = 'VEN' then SetControlText ('GP_NATUREPIECEG', 'DE')
    else  SetControlText ('GP_NATUREPIECEG', 'CF');
    end;
inherited ;
end;

procedure TOF_HistorisePiece.AffecteDispo;
var dQte : double;
    stChamp, stDepot : string;
    iCoef, iIndPiece, iIndLig, iIndNomen : integer;
    TobL, TobP, TobN, TobSerie : TOB;
    TSql : TQuery;
begin
stChamp := '';
iCoef := -1; // on fait l'inverse d'une piece donc -1 pour le plus
if (Pos ('LC', GetInfoParpiece(GetControlText ('GP_NATUREPIECEG'), 'GPP_QTEPLUS')) > 0) then
    stChamp := '_LIVRECLIENT'
else if (Pos ('RC', GetInfoParpiece(GetControlText ('GP_NATUREPIECEG'), 'GPP_QTEPLUS')) > 0) then
    stChamp := '_RESERVECLI'
else if (Pos ('PRE', GetInfoParpiece(GetControlText ('GP_NATUREPIECEG'), 'GPP_QTEPLUS')) > 0) then
    stChamp := '_PREPACLI'
else if (Pos ('LF', GetInfoParpiece(GetControlText ('GP_NATUREPIECEG'), 'GPP_QTEPLUS')) > 0) then
    stChamp := '_LIVREFOU'
else if (Pos ('RF', GetInfoParpiece(GetControlText ('GP_NATUREPIECEG'), 'GPP_QTEPLUS')) > 0) then
    stChamp := '_RESERVEFOU'
else
begin
iCoef := 1; // on fait l'inverse d'une piece donc 1 pour le moins
if (Pos ('LC', GetInfoParpiece(GetControlText ('GP_NATUREPIECEG'), 'GPP_QTEMOINS'))> 0) then
    stChamp := '_LIVRECLIENT'
else if (Pos ('RC', GetInfoParpiece(GetControlText ('GP_NATUREPIECEG'), 'GPP_QTEMOINS')) > 0) then
    stChamp := '_RESERVECLI'
else if (Pos ('PRE', GetInfoParpiece(GetControlText ('GP_NATUREPIECEG'), 'GPP_QTEMOINS')) > 0) then
    stChamp := '_PREPACLI'
else if (Pos ('LF', GetInfoParpiece(GetControlText ('GP_NATUREPIECEG'), 'GPP_QTEMOINS')) > 0) then
    stChamp := '_LIVREFOU'
else if (Pos ('RF', GetInfoParpiece(GetControlText ('GP_NATUREPIECEG'), 'GPP_QTEMOINS')) > 0) then
    stChamp := '_RESERVEFOU';
end;
if stChamp <> '' then
    begin
    TobLigne := Tob.Create ('', nil, -1);
    TobNomenclature := Tob.Create('', nil, -1);
    TobNomenAPlat := Tob.Create('', nil, -1);
    TobSerie := Tob.Create ('', nil, -1);
    for iIndPiece := 0 to TobPiece.Detail.Count - 1 do
        begin
        TobP := TobPiece.Detail [iIndPiece];
        TSql := OpenSql ('SELECT GL_TIERS, GL_NUMLIGNE, GL_NATUREPIECEG, GL_ARTICLE, GL_DEPOT, ' +
                         'GL_QTESTOCK, GL_MTRESTE, GL_TYPELIGNE, GL_REFCATALOGUE, ' +
                         'GL_QUALIFQTEVTE, GL_QUALIFQTEACH, GL_QUALIFQTESTO, ' +
                         'GL_TENUESTOCK, GL_ENCONTREMARQUE, GL_TYPEARTICLE, GL_INDICENOMEN,' +
                         'GL_TYPENOMENC, GL_INDICESERIE, GL_FOURNISSEUR, GL_REFARTSAISIE FROM LIGNE WHERE ' +
                            'GL_NATUREPIECEG="' + TobP.GetValue ('GP_NATUREPIECEG') + '" AND ' +
                            'GL_SOUCHE="' + TobP.GetValue ('GP_SOUCHE') + '" AND GL_NUMERO=' +
                            IntToStr (TobP.GetValue ('GP_NUMERO')) + ' AND GL_INDICEG=' +
                            IntToStr (TobP.GetValue ('GP_INDICEG')), True,-1, '', True);
        if not TSql.Eof then
            begin
            TobLigne.LoadDetailDB ('', '', '', TSql, False);
            end;
        Ferme (TSql);

        if TobLigne.Detail.Count > 0 then
            begin
            // les nomenclatures
            TobNomenclature.ClearDetail;
            ChargeNomenclature (iIndPiece);
            for iIndLig := 0 to TobNomenclature.Detail.Count - 1 do
                begin
                TobL := TobLigne.FindFirst (['GL_ARTICLE'], [TobNomenclature.Detail[iIndLig].GetValue ('ARTICLE')], True);
                if TobL <> nil then
                    begin
                    dQte := TobL.GetValue ('GL_QTERESTE'); {NEWPIECE - DBR}
                    stDepot := TobL.GetValue ('GL_DEPOT');
                    TobL := TobNomenclature.Detail [iIndLig];
                    TobNomenAPlat.ClearDetail;
                    NomenAPlat (TobL, TobNomenAPlat, dQte);
                    for iIndNomen := 0 to TobNomenAPlat.Detail.Count - 1 do
                        begin
                        TobN := TobNomenAPlat.Detail [iIndNomen];
                        if TobN.GetValue ('GLN_TENUESTOCK') = 'X' then
                            begin
                            ModifieTableDispo (TobL, TobN,
                                               TobN.GetValue ('GLN_ARTICLE'),
                                               stChamp, stDepot, iCoef,
                                               TobN.GetValue ('GLN_QTE'));
                            end;
                        end;
                    end;
                end;

            for iIndLig := 0 to TobLigne.Detail.Count - 1 do
                begin
                // les articles
                TobL := TobLigne.Detail[iIndLig];
                if (TobL.GetValue ('GL_TYPELIGNE') = 'ART') and
                   (TobL.GetValue ('GL_TENUESTOCK') = 'X') and
                   (TobL.GetValue ('GL_TYPEARTICLE') <> 'NOM') and
                   (TobL.GetValue ('GL_TYPENOMENC') <> 'ASS') then
                    begin
                    dQte := TobL.GetValue ('GL_QTERESTE'); {NEWPIECE - DBR}
                    ModifieTableDispo (TobL, nil, TobL.GetValue ('GL_ARTICLE'),
                                       stChamp, TobL.GetValue ('GL_DEPOT'),
                                       iCoef, dQte);
                    // les numéros de série
                    if (TobL.GetValue ('GL_TYPELIGNE') = 'ART') and
                       (TobL.GetValue ('GL_INDICESERIE') > 0) then
                        begin
                        TSql := OpenSql ('SELECT * FROM LIGNESERIE WHERE GLS_NATUREPIECEG="' +
                                            TobP.GetValue ('GP_NATUREPIECEG') +
                                            '" AND GLS_SOUCHE="' + TobP.GetValue ('GP_SOUCHE') +
                                            '" AND GLS_NUMERO=' + IntToStr (TobP.GetValue ('GP_NUMERO')) +
                                            ' AND GLS_INDICEG=' + IntToStr (TobP.GetValue ('GP_INDICEG')) +
                                            ' AND GLS_NUMLIGNE=' + IntToStr (TobL.GetValue ('GL_NUMLIGNE')), True,-1, '', True);
                        if not TSql.Eof then
                            begin
                            TobSerie.LoadDetailDb ('', '', '', TSql, False);
                            end;
                        Ferme (TSql);
                        if TobSerie.Detail.Count > 0 then
                            begin
                            if (Pos ('RC', GetInfoParpiece (GetControlText ('GP_NATUREPIECEG'), 'GPP_QTEPLUS')) > 0) then
                                ModifieTableDispoSerie (TobSerie, TobL.GetValue ('GL_DEPOT'),
                                                        'GQS_ENRESERVECLI')
                            else if (Pos ('PRE', GetInfoParpiece (GetControlText ('GP_NATUREPIECEG'), 'GPP_QTEPLUS')) > 0) then
                                ModifieTableDispoSerie (TobSerie, TobL.GetValue ('GL_DEPOT'),
                                                        'GQS_ENPREPACLI');
                            end;
                        end;
                    end;

                // les articles en contremarque
                if (TobL.GetValue ('GL_TYPELIGNE') = 'ART') and
                   (TobL.GetValue ('GL_ENCONTREMARQUE') = 'X') and
                   (pos ('LIVRE', stChamp) = 0) then
                    begin
                    ModifieTableDispoContreM (TobL, stChamp);
                    end;
                end;
            end;
        end;
    TobLigne.Free;
    TobNomenclature.Free;
    TobNomenAPlat.Free;
    TobSerie.Free;
    end;
end;

procedure TOF_HistorisePiece.ChargeLesSousNomen (TobGroupeNomen, TobNomen : TOB;
                                                 LaLig, MaxNiv, iDep : integer);
Var LeNiv, Niv, iInd, Lig : integer ;
    TobLN, TobPere, TobLoc : TOB ;
begin
for LeNiv := 1 to MaxNiv do
    begin
    for iInd := iDep to TobNomen.Detail.Count-1 do
        begin
        TobLN := TobNomen.Detail [iInd] ;
        Niv := TobLN.GetValue ('GLN_NIVEAU');
        Lig := TobLN.GetValue ('GLN_NUMLIGNE');
        if Lig <> LaLig then Continue ;
        if Niv = LeNiv then
           begin
           if Niv = 1 then TobPere := TobGroupeNomen else
              begin
              TobPere := TobGroupeNomen.FindFirst (['GLN_NUMLIGNE', 'GLN_NIVEAU',
                                                    'GLN_ARTICLE', 'GLN_NUMORDRE'],
                                                   [Lig, Niv-1,
                                                    TobLN.GetValue ('GLN_COMPOSE'),
                                                    TobLN.GetValue('GLN_ORDRECOMPO')],
                                                   True) ;
              end ;
           if TobPere <> Nil then
              begin
              TobLoc := Tob.Create ('LIGNENOMEN', TobPere, -1);
              TobLoc.Dupliquer (TobLN, False, True);
              end ;
           end ;
        end ;
    end ;
end;

procedure TOF_HistorisePiece.ChargeNomenclature (iIndPiece : integer);
Var iInd, OldL, Lig, MaxNiv, iIndNomen, Niv, IndiceNomen : integer ;
    TobP, TobL, TobLN, TobNomen, TobGroupeNomen, TobLoc : TOB ;
    OkN : boolean ;
    TSql : TQuery ;
begin
OkN := False;
OldL := -1;
IndiceNomen := 1;
TobP := TobPiece.Detail [iIndPiece];

for iInd := 0 to TobLigne.Detail.Count-1 do
    begin // pour avant tout mettre à zéro l'indice dans la nomenclature
    TobL := TobLigne.Detail[iInd];
    TobL.PutValue ('GL_INDICENOMEN', 0);
    if ((TobL.GetValue ('GL_TYPEARTICLE')='NOM') and (TobL.GetValue('GL_TYPENOMENC')='ASS')) then OkN:=True ;
    end ;
if Not OkN then Exit ;

TobNomen := Tob.Create ('', nil, -1);

TSql := OpenSQL ('SELECT * FROM LIGNENOMEN WHERE GLN_NATUREPIECEG="' +
                 TobP.GetValue ('GP_NATUREPIECEG') + '" AND ' +
                 'GLN_SOUCHE="' + TobP.GetValue ('GP_SOUCHE') + '" AND GLN_NUMERO=' +
                 IntToStr (TobP.GetValue ('GP_NUMERO')) + ' AND GLN_INDICEG=' +
                 IntToStr (TobP.GetValue ('GP_INDICEG')) + ' ORDER BY GLN_NUMLIGNE, GLN_NIVEAU, GLN_NUMORDRE', True,-1, '', True) ;
TobNomen.LoadDetailDB ('', '', '', TSql, True, False) ;
Ferme(TSql) ;

for iInd := 0 to TobNomen.Detail.Count-1 do
    begin
    TobLN := TobNomen.Detail [iInd] ;
    Lig := TobLN.GetValue ('GLN_NUMLIGNE') ;
    if OldL <> Lig then
       begin
       TobGroupeNomen := Tob.Create ('', TobNomenclature, -1);
       MaxNiv := -1 ;
       TobGroupeNomen.AddChampSup ('UTILISE', False);
       TobGroupeNomen.PutValue ('UTILISE', '-');
       TobGroupeNomen.AddChampSup ('ARTICLE', False);
       TobGroupeNomen.PutValue ('ARTICLE', TobLN.GetValue ('GLN_COMPOSE'));
       for iIndNomen := iInd to TobNomen.Detail.Count-1 do
           begin
           TobLoc := TobNomen.Detail [iIndNomen];
           if TobLoc.GetValue ('GLN_NUMLIGNE') <> Lig then Break ;
           Niv := TobLoc.GetValue ('GLN_NIVEAU');
           if Niv > MaxNiv then MaxNiv := Niv;
           end ;
       ChargeLesSousNomen (TobGroupeNomen, TobNomen, Lig, MaxNiv, iInd);
       end ;
    OldL:=Lig ;
    end ;
for iInd := 0 to TobLigne.Detail.Count-1 do
    begin
    TobL := TobLigne.Detail [iInd] ;
    if ((TobL.GetValue('GL_TYPEARTICLE')='NOM') and (TobL.GetValue('GL_TYPENOMENC')='ASS')) then
       begin
       TobL.PutValue ('GL_INDICENOMEN', IndiceNomen);
       Inc (IndiceNomen) ;
       end ;
    end ;
TobNomen.Free;
end;

function TOF_HistorisePiece.ConstruitWhere (stPrefixe : string;
                                            TobPiece : TOB) : string;
var iInd : integer;
    TobP : TOB;
    stWhere : string;
    stCleAna : string;
begin
stWhere := '';
for iInd:=0 to TobPiece.Detail.Count - 1 do
    begin
    TobP := TobPiece.Detail[iInd];
    if stPrefixe <> 'YVA_' then
        begin
        if stWhere <> '' then stWhere := stWhere + ' OR ';
        stWhere := stWhere + '(' + stPrefixe + 'SOUCHE="'  + TobP.GetValue ('GP_SOUCHE') + '"' +
                           ' AND ' + stPrefixe + 'NUMERO=' + IntToStr (TobP.GetValue ('GP_NUMERO')) +
                           ' AND ' + stPrefixe + 'INDICEG=' + IntToStr (TobP.GetValue ('GP_INDICEG')) + ')';
        end else
        begin
        stCleAna := EncodeRefCPGescom (TobP);
        if stWhere = '' then stWhere := 'YVA_IDENTIFIANT in ("'
        else stWhere := stWhere + ', "';
        stWhere := stWhere + stCleAna + '"';
        end;
    end;
if stPrefixe = 'YVA_' then stWhere := stWhere + ')';
Result := stWhere;
end;

procedure TOF_HistorisePiece.Historisation;
var stWhere, stClause : string;
begin
AffecteDispo;

stWhere := ConstruitWhere ('GP_', TobPiece);
if bHistorique then
    begin
    stClause := 'UPDATE PIECE SET GP_VIVANTE="-",GP_DATEMODIF="' + usDateTime(Date) + '"' +
                ' WHERE GP_NATUREPIECEG="' + GetControlText ('GP_NATUREPIECEG') + '"';
    end
    else
    begin
    stClause := 'DELETE FROM PIECE WHERE ' +
                    'GP_NATUREPIECEG="' + GetControlText ('GP_NATUREPIECEG') + '"';
    end;
if stWhere <> '' then stClause := stClause + ' AND (' +
                stWhere + ')';

ExecuteSql (stClause);
stClause := StringReplace (stClause, 'GP_', 'GL_', [rfReplaceAll]);
stClause := StringReplace (stClause, 'PIECE ', 'LIGNE ', [rfReplaceAll]);
ExecuteSql (stClause);
if not bHistorique then
    begin
    stClause := StringReplace (stClause, 'GL_', 'GLL_', [rfReplaceAll]);
    stClause := StringReplace (stClause, 'LIGNE', 'LIGNELOT', [rfReplaceAll]);
    ExecuteSql (stClause); // LIGNELOT
    stClause := StringReplace (stClause, 'GLL_', 'GLN_', [rfReplaceAll]);
    stClause := StringReplace (stClause, 'LIGNELOT', 'LIGNENOMEN', [rfReplaceAll]);
    ExecuteSql (stClause); // LIGNENOMEN
    stClause := StringReplace (stClause, 'GLN_', 'GPA_', [rfReplaceAll]);
    stClause := StringReplace (stClause, 'LIGNENOMEN', 'PIECEADRESSE', [rfReplaceAll]);
    ExecuteSql (stClause); // PIECEADRESSE
    stClause := StringReplace (stClause, 'GPA_', 'GPB_', [rfReplaceAll]);
    stClause := StringReplace (stClause, 'PIECEADRESSE', 'PIEDBASE', [rfReplaceAll]);
    ExecuteSql (stClause); // PIEDBASE
    stClause := StringReplace (stClause, 'GPB_', 'GPE_', [rfReplaceAll]);
    stClause := StringReplace (stClause, 'PIEDBASE', 'PIEDECHE', [rfReplaceAll]);
    ExecuteSql (stClause); // PIEDECHE
    stClause := StringReplace (stClause, 'GPE_', 'GPT_', [rfReplaceAll]);
    stClause := StringReplace (stClause, 'PIEDECHE', 'PIEDPORT', [rfReplaceAll]);
    ExecuteSql (stClause); // PIEDPORT
    if (GetInfoParpiece (GetControlText ('GP_NATUREPIECEG'),
                         'GPP_COMPANALLIGNE') <> 'SAN') or
       (GetInfoParPiece (GetControlText ('GP_NATUREPIECEG'),
                         'GPP_COMPANALPIED') <> 'SAN') then
        begin
        stWhere := ConstruitWhere ('YVA_', TobPiece);
        stClause := 'DELETE FROM VENTANA WHERE ' + stWhere;
        ExecuteSql (stClause); // VENTANA
        end;
    end;
end;

procedure TOF_HistorisePiece.LanceHistorisation ;
var F : TFMul ;
    TobP : TOB;
    TSql : TQuery;
    stWhere : string;
    iInd : integer;
    ioerr : TIOErr ;
begin
TobPiece := Tob.Create ('', Nil, -1);
if TobPiece = nil then exit ;

F:=TFMul(Ecran);
if (F.FListe.NbSelected=0) and (not F.FListe.AllSelected) then
    begin
{$IFDEF EAGLCLIENT}
{$ELSE}
    if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
{$endIF}
    HShowMessage('0;'+F.Caption+';'+TexteMessage[1]+';W;O;O;O;','','') ;
    exit;
    end;

if HShowMessage('0;'+F.Caption+';'+TexteMessage[2]+';E;YN;N;N;',
                '','') = mrYes then
    begin
    if GetInfoParPiece (GetControlText ('GP_NATUREPIECEG'), 'GPP_HISTORIQUE') = 'X' then
        BHistorique := true
    else BHistorique := False;
    if F.FListe.AllSelected then
        begin
        stWhere := RecupWhereCritere(F.Pages);
        TSql := OpenSql ('SELECT GP_NATUREPIECEG, GP_SOUCHE, GP_NUMERO, GP_INDICEG FROM PIECE ' + stWhere, True,-1, '', True);
        if not TSql.Eof then
            begin
            TobPiece.LoadDetailDB ('', '', '', TSql, False);
            end;
        Ferme (TSql);
        end else
        begin
        for iInd:=0 to F.FListe.NbSelected-1 do
            begin
            F.FListe.GotoLeBOOKMARK(iInd);
{$IFDEF EAGLCLIENT}
            F.Q.TQ.Seek(F.FListe.Row-1) ;
{$endIF}
            TobP := Tob.Create ('', TobPiece, -1);
            TobP.AddChampSup ('GP_NATUREPIECEG', False);
            TobP.AddChampSup ('GP_SOUCHE', False);
            TobP.AddChampSup ('GP_NUMERO', False);
            TobP.AddChampSup ('GP_INDICEG', False);
            TobP.PutValue ('GP_NATUREPIECEG', GetControlText ('GP_NATUREPIECEG'));
            TobP.PutValue ('GP_SOUCHE', F.Q.FindField('GP_SOUCHE').AsString);
            TobP.PutValue ('GP_NUMERO', F.Q.FindField('GP_NUMERO').AsInteger);
            TobP.PutValue ('GP_INDICEG', F.Q.FindField('GP_INDICEG').AsInteger);
            end;
        end;

    if TobPiece.Detail.Count > 0 then
        begin
        ioErr := Transactions (Historisation, 1);
        Case ioerr of
            oeUnknown : begin
                        MessageAlerte('Erreur lors de l''historisation des pièces');
                        end;
            end;
        end;
    if F.FListe.AllSelected then F.FListe.AllSelected:=False else F.FListe.ClearSelected;
    F.bSelectAll.Down := False ;
    end;
TobPiece.Free;
end;

procedure TOF_HistorisePiece.ModifieTableDispo (TobL, TobN : TOB;
                                                stArticle, stChamp : string;
                                                stDepot : string;
                                                iCoef : integer; dQte : double);
var dQteDispo : double;
    TSql : TQuery;
    BEof : Boolean;
    stClause : string;
begin
dQteDispo := 0;
stClause := 'SELECT ' + 'GQ' + stChamp + ' FROM DISPO ' +
                ' WHERE ' + 'GQ_ARTICLE="' + stArticle + '" AND ' +
                'GQ_DEPOT="' + stDepot + '" AND ' +
                'GQ_CLOTURE="-"';

TSql := OpenSql (stClause, True,-1, '', True);
BEof := TSql.Eof;
if not BEof then
    begin
    dQteDispo := TSql.Fields [0].AsFloat;
    end;
Ferme (Tsql);
if not BEof then
    begin
    dQte := dQte * iCoef;
    dQte := Arrondi (dQte/GetRatio (TobL, TobN, trsStock), V_PGI.OkDecQ);
    stClause := 'UPDATE DISPO SET ' + 'GQ' + stChamp + '=' +
                     VariantToSQL (dQteDispo + dQte) + ' WHERE ' +
                    'GQ_ARTICLE="' + stArticle + '" AND ' +
                    'GQ_DEPOT="' + stDepot + '" AND ' +
                    'GQ_CLOTURE="-"' ;
    ExecuteSql (stClause);
    end;
end;

procedure TOF_HistorisePiece.ModifieTableDispoContreM (TobL : Tob;
                                                       stChamp : string);
var dQte : double;
    TobDispoContreM, TobDClient, TobDTous : TOB;
    stCodeClient, stCodeFournisseur : string;
    bTousCree, bTousModifie : boolean;
    TSql : TQuery;
    iRang : integer;
begin
bTousCree := False;
bTousModifie := False;
TobDTous := nil;

TobDispoContreM := Tob.Create ('DISPOCONTREM', nil, -1);
if GetInfoParpiece(TobL.GetValue ('GL_NATUREPIECEG'), 'GPP_VENTEACHAT') = 'VEN' then
    begin
    stCodeClient := TobL.GetValue ('GL_TIERS');
    stCodeFournisseur := TobL.GetValue ('GL_FOURNISSEUR');
    end else
    begin
    stCodeClient := TobL.GetValue ('GL_FOURNISSEUR');
    stCodeFournisseur := TobL.GetValue ('GL_TIERS');
    end;


TSql := OpenSql ('SELECT * FROM DISPOCONTREM WHERE GQC_REFERENCE="' +
                     TobL.GetValue ('GL_REFCATALOGUE') + '" AND ' +
                    ' GQC_FOURNISSEUR="' + stCodeFournisseur + '" AND ' +
                    ' (GQC_CLIENT="' + stCodeClient + '" OR GQC_CLIENT="") AND ' +
                    ' GQC_DEPOT="' + TobL.GetValue ('GL_DEPOT') + '"',
                 True,-1, '', True);
TobDispoContreM.LoadDetailDB ('DISPOCONTREM', '', '', TSql, False);

if TobDispoContreM.Detail.Count > 0 then
    begin
    dQte := TobL.GetValue ('GL_QTERESTE'); {NEWPIECE - DBR}
    dQte := Arrondi (dQte/GetRatio (TobL, nil, trsStock), V_PGI.OkDecQ);
    TobDClient := TobDispoContreM.FindFirst (['GQC_CLIENT'], [stCodeClient], false);
    if TobDClient <> nil then
    begin
      TobDTous := TobDispoContreM.FindFirst (['GQC_CLIENT'], [''], false);
      TobDClient.PutValue ('GQC' + stChamp,
                           TobDClient.GetValue ('GQC' + stChamp) - dQte);
      if pos ('CLI', stChamp) > 0 then
          begin
          if TobDClient.GetValue ('GQC_PHYSIQUE') >
              TobDClient.GetValue ('GQC_PREPACLI') + TobDClient.GetValue ('GQC_RESERVECLI') then
              begin
              if TobDTous = nil then
                  begin
                  bTousCree := True;
                  bTousModifie := True;
                  TobDTous := Tob.Create ('DISPOCONTREM', nil, -1);
                  TobDTous.Dupliquer (TobDClient, false, True);
                  TobDTous.PutValue ('GQC_RANG', 0);
                  TobDTous.PutValue ('GQC_CLIENT', '');
                  TobDTous.PutValue ('GQC_PREPACLI', 0);
                  TobDTous.PutValue ('GQC_RESERVECLI', 0);
                  TobDTous.PutValue ('GQC_RESERVEFOU', 0);
                  TobDTous.PutValue ('GQC_PHYSIQUE',
                                      TobDClient.GetValue ('GQC_PHYSIQUE') -
                                          (TobDClient.GetValue ('GQC_PREPACLI') +
                                           TobDClient.GetValue ('GQC_RESERVECLI')));
                  end else
                  begin
                  bTousModifie := True;
                  TobDTous.PutValue ('GQC_PHYSIQUE',
                                      TobDTous.GetValue ('GQC_PHYSIQUE') +
                                          TobDClient.GetValue ('GQC_PHYSIQUE') -
                                          (TobDClient.GetValue ('GQC_PREPACLI') +
                                           TobDClient.GetValue ('GQC_RESERVECLI')));
                  end;
              TobDClient.PutValue ('GQC_PHYSIQUE',
                                      TobDClient.GetValue ('GQC_PREPACLI') +
                                          TobDClient.GetValue ('GQC_RESERVECLI'));
              end;
          end;
      if (TobDClient.GetValue ('GQC_RESERVEFOU') = 0) and
         (TobDClient.GetValue ('GQC_RESERVECLI') = 0) and
         (TobDClient.GetValue ('GQC_PREPACLI') = 0) and
         (TobDClient.GetValue ('GQC_PHYSIQUE') = 0) then
          begin
          TobDClient.DeleteDB;
          end else
          begin
          TobDClient.InsertOrUpdateDB;
          end;
    end;
    if bTousModifie then
        begin
        if bTousCree then
            begin
            TSql := OpenSQL ('SELECT MAX(GQC_RANG) FROM DISPOCONTREM', True,-1, '', True) ;
            if TSql.EOF then iRang := 1 else iRang := TSql.Fields[0].AsInteger + 1 ;
            TobDTous.PutValue ('GQC_RANG', iRang);
            Ferme (TSql);
            end;
        TobDTous.InsertOrUpdateDB;
        end;
    end;
TobDispoContreM.Free;
if bTousCree then TobDTous.Free;
end;

procedure TOF_HistorisePiece.ModifieTableDispoSerie (TobSerie : TOB;
                                                     stDepot,
                                                     stChamp : string);
var stSql, stWhere : string;
    iInd : integer;
begin
stSql := 'UPDATE DISPOSERIE SET ' + stChamp + '="-" WHERE ';
stWhere := '';
for iInd := 0 to TobSerie.Detail.Count - 1 do
    begin
    if stWhere <> '' then stWhere := stWhere + ') OR ('
    else stWhere := '(';
    stWhere := stWhere + 'GQS_ARTICLE="' + TobSerie.Detail[iInd].GetValue ('GLS_ARTICLE') +
                '" AND GQS_IDSERIE="' + TobSerie.Detail[iInd].GetValue ('GLS_IDSERIE') +
                '" AND GQS_DEPOT="' + stDepot + '"';
    end;
if stWhere <> '' then ExecuteSql (stSql + stWhere + ')');
end;

/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLHistorisePiece(parms:array of variant; nb: integer ) ;
var F : TForm ;
    MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_HistorisePiece) then TOF_HistorisePiece(MaTOF).LanceHistorisation
 else exit;
end;

Initialization
registerclasses([TOF_HistorisePiece]);
RegisterAglProc('HistorisePiece',TRUE,1,AGLHistorisePiece);
end.
