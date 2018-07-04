{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 14/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PROPTRANSFFSL ()
Mots clefs ... : TOF;PROPTRANSFFSL
*****************************************************************}
Unit PROPTRANSFFSL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     messages,
     windows,
     mPlayer,
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
     EdtEtat,
     EdtREtat,
{$ELSE}
     UtilEagl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Ent1,
     HMsgBox,
     UTOF,
     SaisieList,
     uTob,
     Facture,
     EntGC,
     uTableFiltre
     ;

type
  TBonDePrepa = class(TFFacture)
  end;

Type
  TOF_PROPTRANSFFSL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    Private
      TF : tTableFiltre;
      lClose : Boolean;
      lAvalider : Boolean;
      Procedure HCODEPTRF_Onexit(Sender : tObject);
      Procedure EMPLACEMENT_Onexit(Sender : tObject);
      Procedure TMP_ARTICLE_OnExit(Sender : tObject);
      procedure MnCreeColis_OnClick(Sender: tObject);
      procedure bSelect_Onclick(Sender : tObject);
      Procedure TMP_QTESAISIE_OnExit(Sender : tObject);
      Procedure TMP_MDP_OnExit(Sender : tObject);
      Procedure BtMDP_OnClick(Sender : tObject);

    public
      lMDP, lExitHCODEPTRF, lExitTMP_ARTICLE : boolean;
  end ;

  Type
    TActionGenereTransfert = Class
    public
      procedure ValideTransfert;
    end;

procedure GenerePieceTransfert(cCodePTRF, User, cLieu : String);
function GenereTEM_TREouTRV(TobTmp : tob; NaturePiece, sWhere: string; var NumeroPiece: integer): boolean;
function GetCodeArticle(cCartSais : String) : String;
procedure ErreurSonore(sFileName: string; Form: tWinControl);

var
  RefDeLaTEM, RefDeLaTRE  : string;
  RefDeLaTRV              : string;
  pCodePTRF, pUser, pLieu : string;

Implementation

uses
  HTB97,
  ed_tools,
  paramsoc,
  wcommuns,
  Math,
  Gerepiece,
  Factcomm,
  TiersUtil,
  UtilArticle,
  {$IFNDEF CCS3}
    UtilConfid,
  {$ENDIF CCS3}
  menus
  ;


{***********A.G.L.***********************************************
Auteur  ...... : Denis Koza
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. : Génération des pieces de transfert inter dépot, TEM ou
Suite ........ : TRV
Mots clefs ... : TRANSFERT, DEPOT, PIECE, TEM, TRV
*****************************************************************}
function GenereTEM_TREouTRV(TobTmp : tob; NaturePiece, sWhere: string; var NumeroPiece: integer): boolean;
var
  i, j                              : integer;
  Tiers                             : string;
  sSql, Annee                       : string;
  DepotOrig, DepotDest              : string;
  TobPiece, TobLigne, TobError, TF  : Tob;
  D1, D2                            : TdateTime;
  Quantieme                         : double;

  function MajDevenirPieceTEM(RefDeLaTEM, RefDeLaTRV: string): integer;
  var
    sSql, DatePiece, NaturePieceG, Souche, Numero, IndiceG : string;
  begin
    DatePiece := ReadTokenSt(RefDeLaTEM);
    DatePiece := UsDateTime(StrToDate(copy(DatePiece,1,2) + '/' + copy(DatePiece,3,2) +  '/' + copy(DatePiece,5,4)));
    NaturePieceG := ReadTokenSt(RefDeLaTEM);
    Souche := ReadTokenSt(RefDeLaTEM);
    Numero := ReadTokenSt(RefDeLaTEM);
    IndiceG := ReadTokenSt(RefDeLaTEM);
    sSql := 'UPDATE PIECE SET GP_DEVENIRPIECE = "'+RefDeLaTRV+'" '+
                        ' WHERE GP_DATEPIECE = "' + DatePiece + '"' +
                        ' AND   GP_NATUREPIECEG = "' + NaturePieceG + '"' +
                        ' AND   GP_SOUCHE = "' + Souche + '"' +
                        ' AND   GP_NUMERO = ' + Numero +
                        ' AND   GP_INDICEG = ' + IndiceG;

    Result := ExecuteSQL(sSql);
  end;

begin
  Result := True;

  // Pour optimiser, si possible, remplir la TOB AVANT appel de cette fonction, ce qui permet de n'éxécuter
  // la requête qu'une fois pour les deux pieces (TEM et TRV par exemple), d'où le test (TobTmp.Detail.Count > 0)
  if (TobTmp.Detail.Count > 0) or  wLoadTobFromSql('_TRANSFERT_', sSql, TobTmp, False) then
  begin

    TobError := Tob.Create('_TEMP_', nil, -1);
    try
      InitMoveProgressForm(nil, TraduireMemoire('Génération des transferts inter-dépôt'), '', TobTmp.detail.count, False, True);

      TobTmp.Detail.Sort('DEPOTDEST');  //Dépôt destinataire
      i := 0;

      while i <= (TobTmp.detail.count -1) do
      begin
        Result := False;
        if      NaturePiece = 'TEM' then
        begin
          DepotOrig := TobTmp.Detail[i].GetValue('DEPOTEMET');
          DepotDest := TobTmp.Detail[i].GetValue('DEPOTDEST');
          if DepotOrig = '' then
            HShowMessage('0;PROBLEME : Le dépôt d''origine n''est pas renseigné. La transaction est refusée.;;O;O;O;','','');
          if DepotDest = '' then
            HShowMessage('0;PROBLEME : Le dépôt de destination n''est pas renseigné. La transaction est refusée.;;O;O;O;','','');
          Result := (DepotOrig <> '') and (DepotDest <> '');
        end
        else if NaturePiece = 'TRV' then
        begin
          DepotOrig := TobTmp.Detail[i].GetValue('DEPOTTRAN');
          DepotDest := TobTmp.Detail[i].GetValue('DEPOTDEST');
          if DepotOrig = '' then
            HShowMessage('0;PROBLEME : Le dépôt de transit n''est pas renseigné. La transaction est refusée.;;O;O;O;','','');
          if DepotDest = '' then
            HShowMessage('0;PROBLEME : Le dépôt de destination n''est pas renseigné. La transaction est refusée.;;O;O;O;','','');
          Result := (DepotOrig <> '') and (DepotDest <> '');
        end
        else if NaturePiece = 'TRE' then
        begin
          DepotOrig := TobTmp.Detail[i].GetValue('DEPOTDEST');
          DepotDest := '';
          if DepotOrig = '' then
            HShowMessage('0;PROBLEME : Le dépôt de destination n''est pas renseigné. La transaction est refusée.;;O;O;O;','','');
          Result := (DepotOrig <> '');
        end;

        Tiers := RechercheTiersDepot(DepotOrig);

        if Result then
        begin
          //Prépare la tob PIECE pour création de la pièce TEM
          //1 pièce par dépôt identique
          TobPiece := Tob.Create('PIECE', nil, -1);
          try
            TobPiece.AddChampSupValeur('NATUREPIECEG',NaturePiece);
            TobPiece.AddChampSupValeur('TIERS',Tiers);
            TobPiece.AddChampSupValeur('AFFAIRE','');
            TobPiece.AddChampSupValeur('ETABLISSEMENT',VH^.EtablisDefaut);
            TobPiece.AddChampSupValeur('DOMAINE','');
            if (NaturePiece = 'TRV') then
              TobPiece.AddChampSupValeur('ETATEXPORT','TRA');  // Sera passé en ENV lors de la validation des pièces de transfert.

            TobPiece.AddChampSupValeur('DEPOT',DepotOrig);
            TobPiece.AddChampSupValeur('DEPOTDEST', DepotDest);
            //Calcul du N° de groupage (Pas sur les commandes de type express)
            if (DepotDest <> '') and (TobTmp.Detail[i].GetValue('LIBRE1')<>'C') and (TobTmp.Detail[i].GetValue('LIBRE1')<>'NC')then
            begin
              Annee := FormatDateTime('yyyy',V_PGI.DateEntree);
              D1 := StrToDateTime(FormatDateTime('dd/mm/yyyy',V_PGI.DateEntree));
              D2 := StrToDateTime(FormatDateTime('dd/mm/yyyy',strtodate('01/01/'+Annee)));
              Quantieme := (D1-D2) + 1;
              TobPiece.AddChampSupValeur('REFINTERNE',DepotDest + FloatToStr(Quantieme));
            end
            else
              TobPiece.AddChampSupValeur('REFINTERNE','');

{V500_010_DKZ Début}
            //if (TobTmp.Detail[i].GetValue('LIBRE1')<> '') then  //Commande express agence pour chainage Bon de prépa en réception transfert
            //  TobPiece.AddChampSupValeur('REFEXTERNE',TobTmp.Detail[i].GetValue('REFEXTERNE'))
            //else
            //  TobPiece.AddChampSupValeur('REFEXTERNE','');
{V500_010_DKZ Fin}
            
{V500_010_DKZ}
            TobPiece.AddChampSupValeur('REFEXTERNE',TobTmp.Detail[i].GetValue('REFEXTERNE'));

            TobPiece.AddChampSupValeur('LIBREPIECE1',TobTmp.Detail[i].GetValue('LIBRE1'));

            while (i <= (TobTmp.detail.count -1))
                  and (DepotOrig = iifV(NaturePiece = 'TEM',TobTmp.Detail[i].GetValue('DEPOTEMET'),
                  iifV(NaturePiece = 'TRV',TobTmp.Detail[i].GetValue('DEPOTTRAN'),{TRE} TobTmp.Detail[i].GetValue('DEPOTDEST')))) do
            begin
              MoveCurProgressForm();
              if TobTmp.Detail[i].GetValue('QUANTITE') > 0 then  //Attention Unité de commande / Unité de stock
              begin
                //Prépare la tob LIGNE
                TobLigne := Tob.Create('LIGNE', TobPiece, -1);

                TobLigne.AddChampSupValeur('CODEARTICLE',TobTmp.Detail[i].GetValue('CODEARTICLE'));
                TobLigne.AddChampSupValeur('ARTICLE',TobTmp.Detail[i].GetValue('ARTICLE'));
                TobLigne.AddChampSupValeur('QTEFACT',TobTmp.Detail[i].GetValue('QUANTITE'));
                TobLigne.AddChampSupValeur('QUALIFQTE',wGetSqlFieldValue('GA_QUALIFUNITESTO', 'ARTICLE', 'GA_ARTICLE="' + TobTmp.Detail[i].GetValue('ARTICLE') + '"'));
                TobLigne.AddChampSupValeur('CIRCUIT','');
                TobLigne.AddChampSupValeur('DEPOT',DepotOrig);
                TobLigne.AddChampSupValeur('DATELIVRAISON',TobTmp.Detail[i].GetValue('DATELIVRAISON'));
                if NaturePiece='TEM' then
                  TobLigne.AddChampSupValeur('PIECEORIGINE','')
                else          //TRV ou TRE
                begin
                  TobLigne.AddChampSupValeur('PIECEORIGINE',RefDeLaTEM);
                  TF := Tob.Create('_TEMP_', TobError, -1);
                  TF.AddChampSupValeur('INDICE',i);
                end;
              end;
              i := i + 1;
            end;

            if TobPiece.detail.count > 0 then
            begin
              if CreatePieceFromTob(TobPiece, nil, nil) then
              begin
                if NaturePiece='TEM' then
                   //RefDeLaTEM := EncodeRefPiece(TobLigne)
                   RefDeLaTEM := EncodeRefPiece(TobPiece.Detail[0]) { JL§ 01/04/2003 Pour éviter un avertissement }
                else
                begin //TRV ou TRE
                   RefDeLaTRV := EncodeRefPiece(Tobpiece);
                   //Mise à jour de GP_DEVENIRPIECE pour le TEM
                   MajDevenirPieceTEM(RefDeLaTEM, RefDeLaTRV);
                   NumeroPiece := TobPiece.getvalue('GP_NUMERO');
                end;
              end
              else
              begin
                Result := False;
                if (NaturePiece = 'TRV') or (NaturePiece = 'TRE') and (assigned(TobError)) and (TobError.detail.count > 0) then
                begin
                  for j := 0 to TobError.detail.count -1 do
                  begin
                  ;
                  end;
                end;
              end;
            end;
          finally
            TobPiece.free;
          end;
        end
        else
        begin
          V_PGI.IoError := oeUnknown;
          break;
        end;  
      end;
      FiniMoveProgressForm;
    finally
      TobError.free;
    end;
    if TobTmp.detail.count > 0 then
      TobTmp.UpdateDB;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Pascal Gautherot
Créé le ...... : 17/03/2003
Modifié le ... : 17/03/2003
Description .. : On considère par défaut que le code saisi est un code
Suite ........ : barre => on va rechercher un article ayant cette valeur
Suite ........ : dans le champs GA_CODEBARRE. Si on trouve, on
Suite ........ : ramène le code article. Dans le cas contraire, on considère
Suite ........ : alors que le code saisi est déjà un code article et on le
Suite ........ : garde tel quel. Si plusieurs articles ont le même code barre,
Suite ........ : on ramène la valeur du premier trouvé.
Mots clefs ... :
*****************************************************************}
Function GetCodeArticle(cCartSais : String) : String;
var
  sReq, sCart : String;
  Q : tQuery;
begin
//  sReq := 'Select GA_CODEARTICLE from ARTICLE where GA_CODEBARRE="'+cCartSais+'"';
  sReq := 'Select GA_ARTICLE from ARTICLE where GA_CODEBARRE="'+cCartSais+'"';
  Q := OpenSql(sReq,True);
  try
    if Q.eof then
    begin
      sCart := cCartSais;
      //Pour permettre la saisie d'un code générique
      if wGetArticleFromCodeArticle(cCartSais) <> cCartSais then
        sCart := wGetArticleFromCodeArticle(cCartSais);
    end
    else
      sCart := Q.FindField('GA_ARTICLE').asString;

  finally
    Ferme(Q);
  end;
  Result := sCart;
end;

procedure ErreurSonore(sFileName: string; Form: tWinControl);
var
  Buffer   : array[0..1023] of Char;
  sWinPath : String;
  SON      : TmediaPlayer;
begin
  GetWindowsDirectory(Buffer,1023);
  SetString(sWinPath, Buffer, StrLen(Buffer));
  if FileExists(sWinPath+'\' + sFileName) then
  begin
    SON := tMediaPlayer.Create(nil);
    try
      SON.AutoOpen := True;
      SON.Parent := Form;
      SON.Visible := False;
      SON.Wait := True;
      SON.FileName := sWinPath + '\' + sFileName;
      SON.Open;
      try
        SON.Wait := True;
        SON.Play;
      finally
        SON.Close;
      end;
    finally
      SON.free;
    end;
  end;
end;

procedure TActionGenereTransfert.ValideTransfert;
begin
  GenerePieceTransfert(pCodePTRF, pUser, pLieu);
end;

procedure TOF_PROPTRANSFFSL.MnCreeColis_OnClick(Sender: tObject);
var
  ActionGenereTransfert : TActionGenereTransfert;
begin
  ActionGenereTransfert := TActionGenereTransfert.Create;
  pCodePTRF := thEdit(GetControl('HCODEPTRF')).text;
  pUser     := V_PGI.User;
  pLieu     := thEdit(GetControl('EMPLACEMENT')).text;
  try
    Transactions(ActionGenereTransfert.ValideTransfert,0);
  finally
    ActionGenereTransfert.free;
  end;

  // Retour a l'écran de sélection
{V500_010_DKZ Début}
  //if GetControl('HCODEPTRF')  <> nil then
  //  SetFocusControl('HCODEPTRF');
{V500_010_DKZ Fin}

{V500_010_DKZ Début}
  if GetControl('HCODEPTRF')  <> nil then
  begin
    if not lMDP then
      PostMessage(TWincontrol(GetControl('HCODEPTRF')).handle, WM_SETFOCUS, 0 , 0 )
    else
      SetFocusControl('HCODEPTRF');
  end;
{V500_010_DKZ Fin}
end;



{***********A.G.L.Privé.*****************************************
Auteur  ...... : Pascal Gautherot
Créé le ...... : 17/03/2003
Modifié le ... : 17/03/2003
Description .. : Procedure de validation du transfert.
Suite ........ : 1) on mets a jours la quantité a transférer comme étant la
Suite ........ : quantité saisie moins la quantité déja transférée
Suite ........ : 2) On lance la procédure de transfert en elle même.
Suite ........ : 3) On mets à jours la quantité transférée comme étant la
Suite ........ : quantité saisie
Mots clefs ... : TRANSFERT, DEPOT
*****************************************************************}
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Procedure en partie similaire a celle de même nom existant dans WPIECETRANSFFSL_TOF
// EN cas de modifs dans celle ci, voir si il ne faut pas faire des modifs du même type de l'autre côté
Procedure GenerePieceTransfert(cCodePTRF, User, cLieu : String);
var
  sReq, sWhere                  : String;
  sReqTob                       : String;
  sTypeCde, sRefExt             : string;       
  St                            : string;       
  DatePiece                     : TdateTime;    
  iNumero, IndiceG, i           : integer;      
  ResGC                         : integer;      
  lOk, lTrouveCde, Ok           : Boolean;      
  Q                             : TQuery;       
  F                             : TBonDePrepa;  
  TobF                          : Tob;          
  TobGTE                        : tob;
  NumeroPiece                   : integer;
  TobQteReste, T1, T2           : tob;
  nQteStock, nQteFact           : Double;
  nQteTotalStock, nQteTotalFact : Double;
  io                            : TiOErr;
begin
  NumeroPiece := 0;
  if existeSql('Select GTL_ARTICLE from PROPTRANSFLIG where (GTL_QTESAISIE-GTL_QTETRANS) > 0 and GTL_CODEPTRF = "'+cCodePTRF+'" and GTL_EMPLACEMENT like "'+cLieu+'%"') then
  begin   // Rien ne sert de lancer les moulinettes si il n'y a rien a traiter... B'en non !
    //1)
    sReq := 'Update PROPTRANSFLIG set GTL_PROPOSITION = (GTL_QTESAISIE-GTL_QTETRANS) WHERE GTL_CODEPTRF = "'+cCodePTRF+'" and GTL_EMPLACEMENT like "'+cLieu+'%"' {and GTL_UTILISATEUR = "'+User+'"'};
    ExecuteSql(sReq);
    sWhere := iifV(cCodePTRF <> '',' GTL_CODEPTRF = "' + cCodePTRF + '"', '');
    sWhere := sWhere + iifV((sWhere <> ''), ' AND ', '') +  iifV(cLieu <> '',' GTL_EMPLACEMENT LIKE "' + cLieu + '%" ', '');
    TobGTE   := Tob.Create('_PROPTRANSFENT_', nil, -1);
    try
      //2)
      sReqTob := 'SELECT GTE_LIBRE1 AS LIBRE1, GTE_PIECEPRECEDENT AS REFEXTERNE  , GTE_DEPOTEMET AS DEPOTEMET, GTL_DEPOTDEST AS DEPOTDEST, GDE_DEPOTTRANSIT AS DEPOTTRAN ' 
        + '      ,GTL_PROPOSITION AS QUANTITE, GTL_ARTICLE AS ARTICLE, GTL_CODEARTICLE AS CODEARTICLE '
        + '      ,GTL_DATEMODIF AS DATELIVRAISON, GTL_QTESAISIE AS QTESAISIE, GTL_QTETRANS AS QTETRANS '  //DKZ à modifier Date GTL_DATEMODIF pas terrible voir pour ajouter zone GTL_DATELIVRAISON
        + ' FROM PROPTRANSFLIG '
        + ' LEFT JOIN PROPTRANSFENT ON (GTL_CODEPTRF = GTE_CODEPTRF) '
        + ' LEFT JOIN DEPOTS ON (GTL_DEPOTDEST = GDE_DEPOT) '
        + iifV(sWhere <> '', ' WHERE ' + sWhere, '')
        + ' ORDER BY GTL_CODEPTRF'
        ;
      if wLoadTobFromSql('_TRANSFERT_', sReqTob, TobGTE, False) then  // Permet de ne charger la TOB qu'une seule fois plutôt que de le faire dans la fonction (donc deux fois)
      begin
       	sTypeCde := TobGTE.detail[0].Getvalue('LIBRE1');
       	sRefExt  := TobGTE.detail[0].Getvalue('REFEXTERNE');
        lOk := GenereTEM_TREouTRV(TobGTE, 'TEM', sWhere, NumeroPiece);
        if lOk then
        begin
          if (sTypeCde='C') or (sTypeCde='NC') then  //Urgente livrée client ou non-urgente livrée client
          begin
            lOk := GenereTEM_TREouTRV(TobGTE, 'TRE', sWhere, NumeroPiece);
            if lOk then
            begin
            iNumero := 0;
            DatePiece := iDate1900;
            indiceG   := 0;
            if ( Copy(sRefExt,1,3) = 'VEN'              ) and
               ( valeurI(copy(sRefExt,4,8)) > 0         ) and
               ( valeurI(copy(sRefExt,4,8)) < 100000000 ) then
              begin
                Q:=OpenSQL('SELECT GP_NUMERO, GP_DATEPIECE, GP_INDICEG FROM PIECE '
                          +' WHERE GP_VIVANTE="X" '
                          +' AND GP_NATUREPIECEG="CC" '
                          +' AND GP_SOUCHE="GCC" '
                          +' AND GP_NUMERO  = ' + copy(sRefExt,4,8)
                          , True);
                try
                  lTrouveCde := False;
                  if not Q.eof then
                  begin
                    iNumero    := Q.findfield('GP_NUMERO').AsInteger;
                    DatePiece  := Q.findfield('GP_DATEPIECE').AsDateTime;
                    IndiceG    := Q.findfield('GP_INDICEG').AsInteger;
                    lTrouveCde := True;
                  end;
                finally
                  Ferme(Q);
                end;

                if lTrouveCde then
                begin
                  //-------------------------------------------------------------
                  //----------- Génération du bon de préparation ----------------
                  //-------------------------------------------------------------
                  F := TBonDePrepa.create(Application);
                  try
                    F.KeyPreview := False;
                    F.GS.Titres.Add('');

                    F.ToutAllouer;

                    St := 'CC;'+DatetoStr(DatePiece)+';GCC;'+IntToStr(iNumero)+';'+IntToStr(IndiceG)+';';

                    if St <> '' then
                    begin
                      StringToCleDoc(St, F.CleDoc);
                      F.CleDoc.Numligne := 0; F.CleDoc.NoPersp := 0;

                      F.Action := taModif;
                      F.NewNature := 'BLC';
                      F.TransfoPiece := True;
                      F.DuplicPiece := False;
                      F.GppReliquat := (GetInfoParPiece('PRE', 'GPP_RELIQUAT') = 'X'); { NEWPIECE }

                      F.ChargeLaPiece(False, False);
                      F.GereVivante;
                      F.EtudieReliquat;
                      F.AppliqueTransfoDuplic(False);
                      F.BloquePiece(F.Action,True);

                      TobQteReste := Tob.Create('_TEMP_', nil, -1);
                      try
                        //TobBP_Origine => 'TRE'
                        //TobPiece_O    => 'CC'
                        //TobPiece      => 'BLC'
                        for i:=0 to F.TOBPiece.detail.count-1 do
                        begin
                          if F.TobPiece_O.detail[i].G('GL_QTERESTE') > 0 then
                          begin
                            TobF := TobGTE.FindFirst(['ARTICLE'],[F.TOBPiece.Detail[i].GetValue('GL_ARTICLE')],False);
                            if TobF <> nil then
                            begin
                              //Clé unique de PROPTRANSFLIG = GTL_CODEPTRF,GTL_ARTICLE,GTL_DEPOTDEST
                              //donc on peut avoir une seule ligne de transfert pour plusieurs lignes dans la commande
                              //d'origine donc on affecte tant qu'il y a du reste à livrer sur les lignes
                              T1 := TobQteReste.FindFirst(['ARTICLE'],[F.TOBPiece.Detail[i].GetValue('GL_ARTICLE')],True);
                              if T1 <> nil then
                              begin
                                nQteTotalStock := T1.G('QTERESTESTOCK');
                                nQteTotalFact  := T1.G('QTERESTEFACT');
                                nQteStock   := Min(nQteTotalStock, F.TOBPiece_O.Detail[i].G('GL_QTERESTE'));
                                nQteFact    := Min(nQteTotalFact , F.TOBPiece_O.Detail[i].G('GL_QTERESTE'));
                                T1.P('QTERESTESTOCK', nQteTotalStock - nQteStock);
                                T1.P('QTERESTEFACT' , nQteTotalFact  - nQteFact);
                              end
                              else
                              begin
                                nQteTotalStock := TOBF.GetValue('QTESAISIE')-TOBF.GetValue('QTETRANS');
                                nQteTotalFact  := nQteTotalStock;
                                nQteStock   := Min(nQteTotalStock, F.TOBPiece_O.Detail[i].G('GL_QTERESTE'));
                                nQteFact    := Min(nQteTotalFact , F.TOBPiece_O.Detail[i].G('GL_QTERESTE'));
                                T2 := Tob.Create('_TEMP_', TobQteReste, -1);
                                T2.AddChampSupValeur('ARTICLE'      ,TobF.G('ARTICLE') );
                                T2.AddChampSupValeur('QTERESTESTOCK',nQteTotalStock - nQteStock);
                                T2.AddChampSupValeur('QTERESTEFACT' ,nQteTotalFact  - nQteFact);
                              end;
                              F.TobPiece.Detail[i].PutValue('GL_QTESTOCK',nQteStock);
                              F.TobPiece.Detail[i].PutValue('GL_QTEFACT' ,nQteFact );
                              F.TobPiece.Detail[i].PutValue('GL_QTERESTE',nQteStock);
                            end
                            else
                            begin
                              F.TobPiece.Detail[i].PutValue('GL_QTESTOCK',0);
                              F.TobPiece.Detail[i].PutValue('GL_QTEFACT' ,0);
                              F.TobPiece.Detail[i].PutValue('GL_QTERESTE',0);
                            end;
                          end;
                        end;
                      finally
                        TobQteReste.free;
                      end;

                      ResGC := GCPieceCorrecte(F.TOBPiece, F.TOBArticles, F.TOBCatalogu);
                      if ResGC > 0 then
                      begin
                        F.HErr.Execute(ResGC,F.Caption,'');
                        Exit;
                      end;
                      Ok := GereEcheancesGC(F.TOBPiece, F.TOBTiers, F.TOBEches, F.TOBAcomptes, F.TOBPIECERG, F.Action, F.DEV, False);
                      if not Ok then Exit;

                      F.ValideNumeroPiece;
                      F.ValideLaPiece;
                      lOk := (V_PGI.IoError = oeOk);
                      if lOk then
                      begin
                        if (GetInfoParPiece(F.NewNature, 'GPP_IMPIMMEDIATE') = 'X') or (GetInfoParPiece(F.NewNature, 'GPP_VALMODELE') = 'X') or (GetInfoParPiece(F.NewNature, 'GPP_IMPETIQ') = 'X') then
                        begin
                          if ((F.Action = taCreat) or (F.Action = taModif) or (F.DuplicPiece) or (F.TransfoPiece)) and
                            not (F.SaisieTypeAvanc) then
                          begin
                            io := Transactions(F.ValideImpression, 1);
                            if io <> oeOk then
                            begin
                              DeflagEdit(F.TOBPiece);
                              MessageAlerte(F.HTitres.Mess[19]);
                            end;
                          end;
                        end;
                      end;
                    end;
                    F.ToutLiberer;
                  finally
                    F.free;
                  end;
                end;
              end;
            end;
          end
          else if (sTypeCde='A') or (sTypeCde='') then    //'A' ou ''     //Urgente livrée agence ou réapprovisionnement agence
          begin
            lOk := GenereTEM_TREouTRV(TobGTE, 'TRV', sWhere, NumeroPiece);
          end;
        end;
        if lOk then
        begin
          // 3)
          sReq := 'Update PROPTRANSFLIG set GTL_QTETRANS=GTL_QTESAISIE WHERE GTL_CODEPTRF = "'+cCodePTRF+'" and GTL_EMPLACEMENT like "'+cLieu+'%"' {and GTL_UTILISATEUR = "'+User+'"'};
          ExecuteSql(sReq);
          if NumeroPiece <> 0 then
          begin
            //Edition Etiquette
            if      (sTypeCde='C') or (sTypeCde='NC') then  //Urgente livrée client ou non-urgente livrée client
            begin
{V500_004 Début}
              V_PGI.NoPrintDialog := True; //not BApercu à paramétrer pour avoir ou non la fenetre de dialogue impression
              LanceEtat('E','LO1','LT2',False,False,False,Nil,'P1.GP_NUMERO='+IntToStr(NumeroPiece)+' AND P1.GP_NATUREPIECEG="TRE"',TraduireMemoire('Etiquettes colis inter-dépôts'),False);
{V500_004 Fin}
            end
            else if (sTypeCde='A') or (sTypeCde='') then
            begin
{V500_004 Début}
              V_PGI.NoPrintDialog := True; //not BApercu à paramétrer pour avoir ou non la fenetre de dialogue impression
              LanceEtat('E','LO1','LT2',False,False,False,Nil,'P1.GP_NUMERO='+IntToStr(NumeroPiece)+' AND P1.GP_NATUREPIECEG="TRV"',TraduireMemoire('Etiquettes colis inter-dépôts'),False);
{V500_004 Fin}
            end;
            //Edition Livraison
            if (sTypeCde <> 'C') and (sTypeCde <> 'NC') then
            begin
{V500_004 Début}
              V_PGI.NoPrintDialog := True; //not BApercu à paramétrer pour avoir ou non la fenetre de dialogue impression
              LanceEtat('E','LO1','LT3',False,False,False,Nil,'P1.GP_NUMERO='+IntToStr(NumeroPiece),TraduireMemoire('Bon de livraison'),False)
{V500_004 Fin}
            end;
          end;
        end;
      end;
    finally
      TobGTE.free;
    end;
  end;
end;

procedure TOF_PROPTRANSFFSL.bSelect_Onclick(Sender : tObject);
var
  cNumero, cGare : String;
begin
  if (GetControl('HCODEPTRF') <> nil) then
    cNumero := ThEdit(GetControl('HCODEPTRF')).Text
  else  // Normalement, ca ne devrais jamais se passer, a moins que quelqu'un sabote la fiche...
    cNumero := '';

  cGare := thEdit(GetControl('EMPLACEMENT')).text;
  // Pour appliquer une sélection sur la table 'fille', on met à jour la propriété WhereTable
  // Puis on désactive puis réactive le tTableFiltre pour que ce where soit prix en compte.
  TF.WhereTable := 'WHERE GTL_EMPLACEMENT like "'+cGare+'%" and GTL_CODEPTRF="'+cNumero+'"';
  TF.Active := False;
  TF.Active := True;
  lAvalider := False;
end;

procedure TOF_PROPTRANSFFSL.BtMDP_Onclick(Sender : tObject); 
begin
  thEdit(GetControl('TMP_MDP')).Visible := True;
  thLabel(GetControl('HLTMP_MDP')).Visible := True;
  if GetControl('TMP_MDP')  <> nil then
    SetFocusControl('TMP_MDP');
end;

procedure TOF_PROPTRANSFFSL.TMP_ARTICLE_OnExit(Sender : tObject);
var
  cCart, cCartSais, cCodePTRF, cGare  : String;
  T                     : Tob;
  I                     : Integer;
  nSais                 : Double;
  ActionGenereTransfert : TActionGenereTransfert;
begin
  lExitHCODEPTRF   := False;
  lExitTMP_ARTICLE := True;
  if not lClose Then
  begin
    cCartSais := thEdit(GetControl('TMP_ARTICLE')).text;
    if (cCartSais <> '') and (cCartSais <> 'FIN') then
      cCart := GetCodeArticle(cCartSais)
    else
    begin
      if cCartSais = '' then
        cCart := ''
      else
        cCart := 'FIN';
    end;
    cCodePTRF := thEdit(GetControl('HCODEPTRF')).text;
    cGare := thEdit(GetControl('EMPLACEMENT')).text;
    if cCart='FIN' then
    begin
      // Validation de la saisie, génération de la piece
      ActionGenereTransfert := TActionGenereTransfert.Create;
      pCodePTRF := cCodePTRF;
      pUser     := V_PGI.User;
      pLieu     := cGare;
      try
        Transactions(ActionGenereTransfert.ValideTransfert,0);
      finally
        ActionGenereTransfert.free;
      end;

      thEdit(GetControl('TMP_ARTICLE')).text := '';
      // Retour a l'écran de sélection
{V500_010_DKZ Début}
      //if GetControl('HCODEPTRF')  <> nil then
      //   PostMessage(TWincontrol(GetControl('HCODEPTRF')).handle, WM_SETFOCUS, 0 , 0 );
      //  SetFocusControl('HCODEPTRF');
      //end;
{V500_010_DKZ Fin}
{V500_010_DKZ Début}
      if GetControl('HCODEPTRF')  <> nil then
      begin
        if not lMDP then
          PostMessage(TWincontrol(GetControl('HCODEPTRF')).handle, WM_SETFOCUS, 0 , 0 )
        else
          SetFocusControl('HCODEPTRF');
      end;
{V500_010_DKZ Fin}
    end
    else if cCart = '' then
    begin
      if thEdit(GetControl('TMP_MDP')).Visible then
         PostMessage(TWincontrol(GetControl('TMP_MDP')).handle, WM_SETFOCUS, 0 , 0 )
      else
         PostMessage(TWincontrol(GetControl('TMP_ARTICLE')).handle, WM_SETFOCUS, 0 , 0 );
    end
    else
    begin // Contrôle du code article, incrémentation du compteur, message si dépassement de quantité
      if not lMDP then
      begin
        lAvalider := True;
        //T := TF.TobFiltre.FindFirst(['GTL_CODEPTRF','GTL_CODEARTICLE'],[cCodePTRF,cCart],false);
        T := TF.TobFiltre.FindFirst(['GTL_CODEPTRF','GTL_ARTICLE'],[cCodePTRF,cCart],false);
        if T <> nil then
        begin
          i := TF.TobFiltre.Detail.indexOf( T )+1;
          TF.SelectRecord( i );
          nSais := TF.GetValue('GTL_QTESAISIE') + 1;

          if nSais > TF.GetValue('GTL_PROPINITIAL') then
          begin
            ErreurSonore(GetParamSoc('SO_ERRFICSON'), TFSaisieList( ecran ));
            HShowMessage('0;ATTENTION : La quantité saisie est supérieure au besoin;;O;O;O;','','');
          end
          else
          begin
            // Mise à jours.
            TF.StartUpdate;
            TF.PutValue( 'GTL_QTESAISIE', nSais );
            TF.PutValue( 'GTL_UTILISATEUR',String(V_PGI.User));
            TF.Post;
            TF.EndUpdate;
          end;
        end
        else
        begin
          ErreurSonore(GetParamSoc('SO_ERRFICSON'), TFSaisieList( ecran ));
          HShowMessage('0;ATTENTION : Cet article ne fait pas partie de cette demande de réappro;;O;O;O;','','');
        end;
        if GetControl('TMP_ARTICLE')   <> nil then thEdit(GetControl('TMP_ARTICLE')).SetFocus;
      end;
    end;
  end;
  if not lMDP then
    thEdit(GetControl('TMP_ARTICLE')).text := '';
end;

procedure TOF_PROPTRANSFFSL.TMP_QTESAISIE_OnExit(Sender : tObject); 
var
  cCart, cCartSais, cCodePTRF, cGare : String;
  T                     : Tob;
  I                     : Integer;
  nSais, nSaisQte       : Double;
begin
  if not lClose Then
  begin
    cCartSais := thEdit(GetControl('TMP_ARTICLE')).text;
    cCart := GetCodeArticle(cCartSais);
    cCodePTRF := thEdit(GetControl('HCODEPTRF')).text;
    cGare := thEdit(GetControl('EMPLACEMENT')).text;
    nSaisQte := thNumEdit(GetControl('TMP_QTESAISIE')).Value;
    if (lMDP) and (cCart <> '') and (nSaisQte <> 0) then
    begin
      lAvalider := True;
      T := TF.TobFiltre.FindFirst(['GTL_CODEPTRF','GTL_ARTICLE'],[cCodePTRF,cCart],false);
      if T <> nil then
      begin
        i := TF.TobFiltre.Detail.indexOf( T )+1;
        TF.SelectRecord( i );
        nSais := TF.GetValue('GTL_QTESAISIE') + nSaisQte;
        if nSais > TF.GetValue('GTL_PROPINITIAL') then
        begin
          ErreurSonore(GetParamSoc('SO_ERRFICSON'), TFSaisieList( ecran ));
          HShowMessage('0;ATTENTION : La quantité saisie est supérieure au besoin;;O;O;O;','','');
        end
        else
        begin
          // Mise à jours.
          TF.StartUpdate;
          TF.PutValue( 'GTL_QTESAISIE', nSais );
          TF.PutValue( 'GTL_UTILISATEUR',String(V_PGI.User));
          TF.Post;
          TF.EndUpdate;
        end;
      end
      else
      begin
        ErreurSonore(GetParamSoc('SO_ERRFICSON'), TFSaisieList( ecran ));
        HShowMessage('0;ATTENTION : Cet article ne fait pas partie de cette demande de réappro;;O;O;O;','','');
      end;
      if GetControl('TMP_ARTICLE')   <> nil then thEdit(GetControl('TMP_ARTICLE')).SetFocus;
    end;
  end;
  thEdit(GetControl('TMP_ARTICLE')).text := '';
  thNumEdit(GetControl('TMP_QTESAISIE')).Value := 0;
end;

procedure TOF_PROPTRANSFFSL.TMP_MDP_OnExit(Sender : tObject);
begin
  if ValeurI(thEdit(GetControl('TMP_MDP')).Text) = GetParamSoc('SO_MDPSCANCOLIS') then
  begin
    thNumEdit(GetControl('TMP_QTESAISIE')).Visible := not thNumEdit(GetControl('TMP_QTESAISIE')).Visible;
    thLabel(GetControl('HLTMP_QTESAISIE')).Visible := not thLabel(GetControl('HLTMP_QTESAISIE')).Visible;
    lMDP := not lMDP;
  end;
  thEdit(GetControl('TMP_MDP')).Text := '';
  thLabel(GetControl('HLTMP_MDP')).Visible := False;
  thEdit(GetControl('TMP_MDP')).Visible := False;
  if      lExitHCODEPTRF then
     //PostMessage(TWincontrol(GetControl('HCODEPTRF')).handle, WM_SETFOCUS, 0 , 0 )
    thEdit(GetControl('HCODEPTRF')).SetFocus
  else if lExitTMP_ARTICLE then
     //PostMessage(TWincontrol(GetControl('TMP_ARTICLE')).handle, WM_SETFOCUS, 0 , 0 );
    thEdit(GetControl('TMP_ARTICLE')).SetFocus;
end;

Procedure TOF_PROPTRANSFFSL.HCODEPTRF_Onexit(Sender : tObject);
var
  cCodePTRF : string;
  Q         : TQuery;
begin
  lExitHCODEPTRF   := True;
  lExitTMP_ARTICLE := False;
  cCodePTRF := thEdit(GetControl('HCODEPTRF')).text;
  if (not lClose) and (cCodePTRF <> '') then
  begin
    Q:=OpenSQL('SELECT null FROM PROPTRANSFENT '
              +' WHERE GTE_CODEPTRF = "' + cCodePTRF + '"'
              , True);
    try
      if Q.eof then
      begin
        PgiInfo(TraduireMemoire('Le code transfert n''existe pas.'));
        thEdit(GetControl('HCODEPTRF')).SetFocus;
      end;
    finally
      Ferme(Q);
    end;
  end;
end;

Procedure TOF_PROPTRANSFFSL.EMPLACEMENT_Onexit(Sender : tObject);
begin
  if not lClose then
  begin
    bSelect_Onclick(Sender);
    if GetControl('TMP_ARTICLE')   <> nil then thEdit(GetControl('TMP_ARTICLE')).SetFocus;
  end;
end;

procedure TOF_PROPTRANSFFSL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PROPTRANSFFSL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PROPTRANSFFSL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PROPTRANSFFSL.OnLoad ;
begin
  Inherited ;
  AppliquerConfidentialite(TFSaisieList( ecran ),'GTE');
end ;

procedure TOF_PROPTRANSFFSL.OnArgument (S : String ) ;
var
  Q                    : TQuery;
  ListeUserGrp, UserGrp : string;
begin
  Inherited ;
  TF := TFSaisieList( ecran ).LeFiltre;
  lClose := False;
  lAvalider := False;
  // C'est un écran sur lequel on agis sur des lignes existantes. pas question de créer, modifier ou dupliquer
  tToolbarButton97(GetControl('bDelete')).Visible := False;
  tToolbarButton97(GetControl('bInsert')).Visible := False;
  tToolbarButton97(GetControl('bDupliquer')).Visible := False;
  thLabel(GetControl('HLTMP_MDP')).Visible := False;
  thEdit(GetControl('TMP_MDP')).Visible := False;          
  //
  if GetControl('HCODEPTRF')    <> nil then thEdit(GetControl('HCODEPTRF')).Onexit := HCODEPTRF_Onexit;
  if GetControl('EMPLACEMENT')  <> nil then thEdit(GetControl('EMPLACEMENT')).Onexit := EMPLACEMENT_Onexit;
  if GetControl('TMP_ARTICLE')  <> nil then thEdit(GetControl('TMP_ARTICLE')).OnExit := TMP_ARTICLE_OnExit;
  if GetControl('TMP_QTESAISIE')<> nil then thNumEdit(GetControl('TMP_QTESAISIE')).OnExit := TMP_QTESAISIE_OnExit;  
  if GetControl('TMP_MDP')      <> nil then thEdit(GetControl('TMP_MDP')).OnExit := TMP_MDP_OnExit;                 
  if GetControl('BtMDP')        <> nil then tButton(GetControl('BtMDP')).Onclick := BtMDP_Onclick;                  
  if GetControl('MNCREECOLIS')  <> nil then TMenuItem(GetControl('MNCREECOLIS')).OnClick := MnCreeColis_OnClick;
  if GetControl('bSelect')      <> nil then tButton(GetControl('bSelect')).Onclick := bSelect_Onclick;

  //Permet de rentrer en mode saisie Article, Qté sans Scan
  Q:=OpenSQL('SELECT UG_GROUPE FROM USERGRP '
            +' WHERE UG_NUMERO = ' + inttostr(V_PGI.UserGrp)
            , True);
  try
    UserGrp := '';
    if not Q.eof then
      UserGrp := Q.findfield('UG_GROUPE').AsString;
  finally
    Ferme(Q);
  end;

  ListeUserGrp := GetParamsoc('SO_SCANUSERGRP');
  lMDP := not (pos(UserGrp, ListeUserGrp) > 0);
  thNumEdit(GetControl('TMP_QTESAISIE')).Visible := lMDP;
  thLabel(GetControl('HLTMP_QTESAISIE')).Visible := lMDP;

  bSelect_Onclick(Self);
  thEdit(GetControl('HCODEPTRF')).SetFocus;
end ;

procedure TOF_PROPTRANSFFSL.OnClose ;
var
  sReq : String;
  cGare : String;
begin
  Inherited ;
  lClose := True;
  if lAvalider then
  begin
    cGare := thEdit(GetControl('EMPLACEMENT')).text;
    // Il y a eu des lignes saisies, mais on ne les a pas validé => On réinitialise les quantités saisies à la quantité préalablement validée
    sReq := 'Update PROPTRANSFLIG set GTL_QTESAISIE = GTL_QTETRANS WHERE GTL_CODEPTRF = "'+thEdit(GetControl('HCODEPTRF')).text+'" and GTL_EMPLACEMENT like "'+cGare+'%"' { and GTL_UTILISATEUR = "'+V_PGI.User+'"'};
    ExecuteSql(sReq);
  end;
end ;

procedure TOF_PROPTRANSFFSL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PROPTRANSFFSL.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_PROPTRANSFFSL ] ) ;
end.
