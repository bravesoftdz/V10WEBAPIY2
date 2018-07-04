{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PIECETRANSFFSL ()
Mots clefs ... : TOF;PIECETRANSFFSL; 
*****************************************************************}
Unit PIECETRANSFFSL_TOF ;

interface

uses StdCtrls,
     Controls,
     Classes,
     messages,
     windows,
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
     HMsgBox,
     uTOF,
     SaisieList,
     uTob,
     Facture,
     uTableFiltre ;

type
  TBonDePrepa = class(TFFacture)
  end;

type
  TOF_PIECETRANSFFSL = Class (TOF)
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
      procedure TMP_ARTICLE_OnExit(Sender : tObject);
      procedure HNUMERO_OnExit(Sender : tObject);
      procedure MnCreeColis_OnClick(Sender: tObject);
      procedure bSelect_Onclick(Sender : tObject);
  end ;

type
  TActionGenereTransfert = Class
  public
    procedure ValideTransfert;
  end;

procedure GenerePieceTransfert(iNumero : Integer; sDepotEmet : string);

var
  pNumero    : integer;
  pDepotEmet : string;

implementation

uses
  HTB97,
  wCommuns,
  menus,
  PropTransfFSL_TOF,
  factcomm,
  factutil,
  Paramsoc;

procedure TActionGenereTransfert.ValideTransfert;
begin
  GenerePieceTransfert(pNumero, pDepotEmet);
end;

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Procedure en partie similaire a celle de même nom existant dans WPROPTRANSFSL_TOF
// En cas de modifs dans celle ci, voir si il ne faut pas faire des modifs du même type de l'autre côté
procedure GenerePieceTransfert(iNumero : Integer; sDepotEmet : String);
var
  sReq, sWhere: String;
  sReqTob     : String;
  TobGTE      : tob;
  sTypeCde    : String;
  lOk         : boolean;
  NumeroPiece : integer;
  F           : TBonDePrepa;
  St          : string;
  i, ResGC    : integer;
  Ok          : boolean;
  TobF        : tob;
begin
  sWhere := '(GL_QTEPREVAVANC-GL_QTESIT) > 0 and GL_NATUREPIECEG = "CC" AND GL_NUMERO='+IntToStr(iNumero)+' AND GL_VIVANTE = "'+wTrue+'"';
  if existeSql('Select GL_ARTICLE from LIGNE'
                  + ' LEFT JOIN PIECE ON (GL_NUMERO = GP_NUMERO) and (GL_NATUREPIECEG = GP_NATUREPIECEG) AND (GL_SOUCHE = GP_SOUCHE) AND (GL_INDICEG = GP_INDICEG)'
                  + ' WHERE '+sWhere) then
  begin   // Rien ne sert de lancer les moulinettes si il n'y a rien a traiter... B'en non !
    TobGTE   := Tob.Create('_PROPTRANSFENT_', nil, -1);
    try
      // 1) on génère les pièces
      sReqTob := 'SELECT GP_LIBREPIECE1 AS LIBREPIECE1, "" AS REFEXTERNE, "'+sDepotEmet+'" AS DEPOTEMET, GL_DEPOT AS DEPOTDEST, GDE_DEPOTTRANSIT AS DEPOTTRAN '
        + '      ,(GL_QTEPREVAVANC-GL_QTESIT) AS QUANTITE, GL_ARTICLE AS ARTICLE, GL_CODEARTICLE AS CODEARTICLE '
        + '      ,GL_DATEMODIF AS DATELIVRAISON, GL_QTEPREVAVANC AS QTESAISIE, GL_QTESIT AS QTETRANS '  //DKZ à modifier Date GTL_DATEMODIF pas terrible voir pour ajouter zone GTL_DATELIVRAISON
        + '      ,GP_NATUREPIECEG AS NATUREPIECEG, GP_DATEPIECE AS DATEPIECE, GP_INDICEG AS INDICEG, GP_NUMERO AS NUMERO '
        + ' FROM LIGNE '
        + ' LEFT JOIN PIECE ON (GL_NUMERO = GP_NUMERO) and (GL_NATUREPIECEG = GP_NATUREPIECEG) AND (GL_SOUCHE = GP_SOUCHE) AND (GL_INDICEG = GP_INDICEG)'
        + ' LEFT JOIN DEPOTS ON (GL_DEPOT = GDE_DEPOT) '
        + iif(sWhere <> '', ' WHERE ' + sWhere, '')
        + ' ORDER BY GL_NUMERO'
        ;
      if wLoadTobFromSql('_TRANSFERT_', sReqTob, TobGTE, False) then  // Permet de ne charger la TOB qu'une seule fois plutôt que de le faire dans la fonction (donc deux fois)
      begin
        // Faire ici, test selon une valeur d'un champs de PIECE a définir, faire soit TEM+TRV, soit TEM+TRE
        // Selon si c'est expédié au dépot ou directement chez le client.
        NumeroPiece := 0;
       	//sTypeCde := WGetSqlFieldValue('GP_LIBREPIECE1', 'PIECE', 'GP_NATUREPIECEG = "CC" AND GP_NUMERO='+IntToStr(iNumero)+' AND GP_VIVANTE = "'+wTrue+'"');
       	sTypeCde := TobGTE.detail[0].Getvalue('LIBREPIECE1');
        lOk := GenereTEM_TREouTRV(TobGTE, 'TEM', sWhere, NumeroPiece);
        if lOk then
        begin
          if sTypeCde='C' then  //Commande express livrée chez le client
          begin
            lOk := GenereTEM_TREouTRV(TobGTE, 'TRE', sWhere, NumeroPiece);
            //-------------------------------------------------------------
            //----------- Génération du bon de préparation ----------------
            //-------------------------------------------------------------
            F := TBonDePrepa.create(Application);
            try
              F.KeyPreview := False;
              F.GS.Titres.Add('');

              F.ToutAllouer;

              St := 'CC;'+DatetoStr(TobGTE.detail[0].Getvalue('DATEPIECE'))+';GCC;'+IntToStr(iNumero)+';'+IntToStr(TobGTE.detail[0].Getvalue('INDICEG'))+';';

              if St <> '' then
              begin
                StringToCleDoc(St, F.CleDoc);
                F.CleDoc.Numligne := 0; F.CleDoc.NoPersp := 0;

                F.Action := taModif;
                F.NewNature := 'PRE';
                F.TransfoPiece := True;
                F.DuplicPiece := False;

                F.ChargeLaPiece(False);
                F.GereVivante;
                F.EtudieReliquat;
                F.AppliqueTransfoDuplic(False);
                F.BloquePiece(F.Action,True);

                for i:=0 to F.TOBPiece.detail.count-1 do
                begin
                  TobF := TobGTE.FindFirst(['ARTICLE'],[F.TOBPiece.Detail[i].GetValue('GL_ARTICLE')],False);
                  if TobF <> nil then
                  begin
                    F.TobPiece.Detail[i].PutValue('GL_QTESTOCK',TOBF.GetValue('QTESAISIE')-TOBF.GetValue('QTETRANS'));
                    F.TobPiece.Detail[i].PutValue('GL_QTEFACT' ,TOBF.GetValue('QTESAISIE')-TOBF.GetValue('QTETRANS'));
                  end
                  else
                  begin
                    F.TobPiece.Detail[i].PutValue('GL_QTESTOCK',0);
                    F.TobPiece.Detail[i].PutValue('GL_QTEFACT' ,0);
                  end;
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
              end;
              F.ToutLiberer;
            finally
              F.free;
            end;
          end
          else      //'A'      //Commande express livrée en agence
            lOk := GenereTEM_TREouTRV(TobGTE, 'TRV', sWhere, NumeroPiece);
        end;
        if lOk then
        begin
          // 2) si les pièces ont bien été générées, on incrémente la quantité transférée.
          sReq := 'Update LIGNE set GL_QTESIT=GL_QTEPREVAVANC WHERE '+sWhere;
          ExecuteSql(sReq);
          if (NumeroPiece <> 0) then //and (sTypeCde = 'A') then //Editions uniquement si Express 'A'gence
          begin
            //{$IFDEF AGL560}
            if LanceEtat('E','LO1','LT2',true,False,False,Nil,'P1.GP_NUMERO='+IntToStr(NumeroPiece),TraduireMemoire('Etiquettes colis inter-dépôts'),False) < 0 then
            //{$ELSE}
            //if not LanceEtat('E','LO1','LT2',true,False,False,Nil,'P1.GP_NUMERO='+IntToStr(NumeroPiece),TraduireMemoire('Etiquettes colis inter-dépôts'),False) then
            //{$ENDIF}
              V_PGI.IoError:=oeUnknown;
            //{$IFDEF AGL560}
            if LanceEtat('E','LO1','LT3',true,False,False,Nil,'P1.GP_NUMERO='+IntToStr(NumeroPiece),TraduireMemoire('Bon de livraison'),False) < 0 then
            //{$ELSE}
            //if not LanceEtat('E','LO1','LT3',true,False,False,Nil,'P1.GP_NUMERO='+IntToStr(NumeroPiece),TraduireMemoire('Bon de livraison'),False) then
            //{$ENDIF}
              V_PGI.IoError:=oeUnknown;
          end;
        end;
      End;
    finally
      TobGTE.free;
    end;
  end;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////


Procedure TOF_PIECETRANSFFSL.bSelect_Onclick(Sender : tObject);
var
  sdepotEmet : String;
  iNumero : Integer;
Begin
  sdepotEmet := thedit(GetControl('DEPOTEMET')).Text;
  iNumero := ValeurI(thEdit(GetControl('HNUMERO')).Text);
//  if (sDepotEmet <> '') and ExisteSql('Select GDE_DEPOT from DEPOTS Where GDE_DEPOT = "'+sDepotEmet+'"') then
//  begin
    TF.WhereTable := 'WHERE GL_NATUREPIECEG = "CC" AND GL_NUMERO="'+IntToStr(iNumero)+'" AND GL_VIVANTE = "' + wTrue + '"';
    TF.Active := False;
    TF.Active := True;
//  end
//  else
//  begin
//    HShowMessage('0;ATTENTION ! Ce dépôt n''existe pas;;O;O;O;','','');
//    abort;
//  end;
end;

procedure TOF_PIECETRANSFFSL.MnCreeColis_OnClick(Sender: tObject);
var
  ActionGenereTransfert : TActionGenereTransfert;
begin
  ActionGenereTransfert := TActionGenereTransfert.Create;
  pNumero     := ValeurI(thEdit(GetControl('HNUMERO')).Text);
  pDepotEmet  := thEdit(GetControl('DEPOTEMET')).Text;
  try
    Transactions(ActionGenereTransfert.ValideTransfert,1);
  finally
    ActionGenereTransfert.free;
  end;

  // Retour a l'écran de sélection
  if (GetControl('DEPOTEMET')  <> nil) and thEdit(GetControl('DEPOTEMET')).Visible then
    PostMessage(TWincontrol(GetControl('DEPOTEMET')).handle, WM_SETFOCUS, 0 , 0 )
    //thedit(GetControl('DEPOTEMET')).SetFocus
  else
    PostMessage(TWincontrol(GetControl('HNUMERO')).handle, WM_SETFOCUS, 0 , 0 );
    //thEdit(GetControl('HNUMERO')).SetFocus;
end;


procedure TOF_PIECETRANSFFSL.HNUMERO_OnExit(Sender : tObject);
begin
  if not lClose then
  begin
    bSelect_Onclick(Sender);
    if GetControl('TMP_ARTICLE')   <> nil then thEdit(GetControl('TMP_ARTICLE')).SetFocus;
  end;  
end;


procedure TOF_PIECETRANSFFSL.TMP_ARTICLE_OnExit(Sender : tObject);
var
  cCart, cCartSais, sDepotEmet : String;
  T : Tob;
  I, iNumero : Integer;
  nSais : Double;
  ActionGenereTransfert : TActionGenereTransfert;
begin
  if not lClose then
  begin
    cCartSais := thEdit(GetControl('TMP_ARTICLE')).text;
    sDepotEmet := thEdit(GetControl('DEPOTEMET')).text;
    cCart := GetCodeArticle(cCartSais);
    iNumero := ValeurI(thEdit(GetControl('HNUMERO')).text);
    if cCart <> '' then
    begin
      if cCart='FIN' then
      begin
        // Validation de la saisie, génération de la piece
        ActionGenereTransfert := TActionGenereTransfert.Create;
        pNumero     := iNumero;
        pDepotEmet  := sDepotEmet;
        try
          Transactions(ActionGenereTransfert.ValideTransfert,1);
        finally
          ActionGenereTransfert.free;
        end;

        if (GetControl('DEPOTEMET')  <> nil) and thEdit(GetControl('DEPOTEMET')).Visible then
          PostMessage(TWincontrol(GetControl('DEPOTEMET')).handle, WM_SETFOCUS, 0 , 0 )
          //thedit(GetControl('DEPOTEMET')).SetFocus
        else
          PostMessage(TWincontrol(GetControl('HNUMERO')).handle, WM_SETFOCUS, 0 , 0 )
          //thEdit(GetControl('HNUMERO')).SetFocus;
      end
      else
      begin // Contrôle du code article, incrémentation du compteur, message si dépassement de quantité
        lAvalider := True;
        //T := TF.TobFiltre.FindFirst(['GL_NUMERO','GL_NATUREPIECEG','GL_CODEARTICLE'],[iNumero,'CC',cCart],false);
        T := TF.TobFiltre.FindFirst(['GL_NUMERO','GL_NATUREPIECEG','GL_ARTICLE'],[iNumero,'CC',cCart],false);
        if T <> nil then
        begin
          i := TF.TobFiltre.Detail.indexOf( T )+1;
          TF.SelectRecord( i );
          nSais := TF.GetValue('GL_QTEPREVAVANC') + 1;
          if nSais > TF.GetValue('GL_QTESTOCK') then
          begin
            ErreurSonore(GetParamSoc('SO_ERRFICSON'), TFSaisieList( ecran ));
            HShowMessage('0;ATTENTION ! La quantité saisie est supérieure au besoin;;O;O;O;','','');
          end
          else
          begin  // Mise à jours.
            TF.StartUpdate;
            TF.PutValue( 'GL_QTEPREVAVANC', nSais );
            TF.PutValue( 'GL_UTILISATEUR',String(V_PGI.User));
            TF.Post;
            TF.EndUpdate;
          end;
        end
        else
        begin
          ErreurSonore(GetParamSoc('SO_ERRFICSON'), TFSaisieList( ecran ));
          HShowMessage('0;ATTENTION ! Cet article ne fait pas partie de cette demande de réappro;;O;O;O;','','');
        end;
        if GetControl('TMP_ARTICLE')   <> nil then thEdit(GetControl('TMP_ARTICLE')).SetFocus;
      end;
    end;
  end;
  thEdit(GetControl('TMP_ARTICLE')).text := '';
end;


procedure TOF_PIECETRANSFFSL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PIECETRANSFFSL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PIECETRANSFFSL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PIECETRANSFFSL.OnLoad ;
begin
  Inherited ;
//  SetControlText('XX_WHERE',GetControlText('XX_WHERE')+'AND GP_NATUREPIECEG = "CC" AND GP_VIVANTE="'+wTrue+'"');
  if (GetControl('DEPOTEMET')  <> nil) and thEdit(GetControl('DEPOTEMET')).Visible then
    thEdit(GetControl('DEPOTEMET')).SetFocus
  else
    thEdit(GetControl('HNUMERO')).SetFocus;
end ;

procedure TOF_PIECETRANSFFSL.OnArgument (S : String ) ;
begin
  Inherited ;
  TF := TFSaisieList( ecran ).LeFiltre;
  lClose := False;
  lAvalider := False;
  // C'est un écran sur lequel on agis sur des lignes existantes. pas question de créer, modifier ou dupliquer
  tToolbarButton97(GetControl('bDelete')).Visible := False;
  tToolbarButton97(GetControl('bInsert')).Visible := False;
  tToolbarButton97(GetControl('bDupliquer')).Visible := False;
  //
  if GetControl('HNUMERO') <> nil then thEdit(GetControl('HNUMERO')).Onexit := HNUMERO_Onexit;
  if GetControl('TMP_ARTICLE')   <> nil then thEdit(GetControl('TMP_ARTICLE')).OnExit := TMP_ARTICLE_OnExit;
  if GetControl('MNCREECOLIS')  <> nil then TMenuItem(GetControl('MNCREECOLIS')).OnClick := MnCreeColis_OnClick;
  if GetControl('bSelect')     <> nil then tButton(GetControl('bSelect')).Onclick := bSelect_Onclick;
  thEdit(GetControl('DEPOTEMET')).text := GetParamSoc('SO_DEPOTPRINCIPAL');
  if (thEdit(GetControl('DEPOTEMET')).text = '') then
  begin
    ThLabel(GetControl('TDEPOTEMET')).Visible := True;
    thEdit(GetControl('DEPOTEMET')).visible := True;
  end
  else
  begin
    ThLabel(GetControl('TDEPOTEMET')).Visible := False;
    thEdit(GetControl('DEPOTEMET')).Visible := False;
  end;
  bSelect_Onclick(Self);
  if (GetControl('DEPOTEMET')  <> nil) and thEdit(GetControl('DEPOTEMET')).Visible then
    thedit(GetControl('DEPOTEMET')).SetFocus
  else
    if GetControl('HNUMERO')   <> nil then thEdit(GetControl('HNUMERO')).SetFocus;
end ;

procedure TOF_PIECETRANSFFSL.OnClose ;
var
  sWhere, sreq : String;
begin
  Inherited ;
  lClose := True;
  if lAvalider then
  begin
    // Il y a eu des lignes de saisies, mais on ne les a pas validé => On réinitialise les quantité saisies à la quantité préalablement validée
    sWhere := '(GL_QTEPREVAVANC-GL_QTESIT) > 0 and GL_NATUREPIECEG = "CC" AND GL_NUMERO='+IntToStr(ValeurI(thEdit(GetControl('HNUMERO')).Text))+' AND GL_VIVANTE = "'+wTrue+'"';
    sReq := 'Update LIGNE set GL_QTEPREVAVANC=GL_QTESIT WHERE '+sWhere;
    ExecuteSql(sReq);
  end;
end ;

procedure TOF_PIECETRANSFFSL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PIECETRANSFFSL.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_PIECETRANSFFSL ] ) ;
end.
