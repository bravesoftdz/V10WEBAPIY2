{***********UNITE*************************************************
Auteur  ...... : Julien D
Créé le ...... : 19/03/2002                                                      
Modifié le ... : 20/03/2002
Description .. : Source TOF de la FICHE : MBOEXPORTPIECE ()
Mots clefs ... : TOF;MBOEXPORTPIECE
*****************************************************************}
Unit DuplicPIECE_TOF ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
     eMul,MaineAGL,
{$ELSE}
     Fiche, HDB, mul, DBGrids, db,dbTables,Fe_Main,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,
     HStatus,HQry,UTob,FileCtrl,Paramsoc,M3FP ;

Type
  TOF_MBORECALCULPIECE = Class (TOF)
    Private
      stWhere, CodeTiers, CodeTiersOld : string;
      Tob_Tiers       : TOB;
      procedure BatchDuplic;
      procedure SetDuplic;
      procedure SetAllDuplic;
      procedure DuplicPiece(var Tob_Piece : TOB);
      function  ExistePiece(Nature,Souche : string; Numero:integer): boolean;
    Public
      procedure OnLoad                   ; override ;
  end ;

Implementation

procedure TOF_MBORECALCULPIECE.BatchDuplic;
var i : integer;
begin
  // Libération de la TOB Tiers
  Tob_Tiers    := nil ;
  CodeTiersOld := ''; CodeTiers := '';

  with TFMul(Ecran) do
   begin
   if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
     begin PGIInfo('Veuillez sélectionner les pièces à dupliquer', Caption); exit; end;
   SourisSablier;
   try
     if FListe.AllSelected then
     begin
       if PGIAsk('Voulez-vous dupliquer toutes les pièces ?', Caption) <> mrYes then exit;
       if Transactions(SetallDuplic,0) <> oeOK then PGIError('Impossible de dupliquer toutes les pièces', Caption);
       FListe.AllSelected := false;
     end else
     begin
       if PGIAsk('Voulez-vous dupliquer les pièces sélectionnées ?', Caption) <> mrYes then exit;
       InitMove(FListe.NbSelected,'');
       For i := 0 to FListe.NbSelected-1 do
       begin
         FListe.GotoLeBookMark(i);
         {$IFDEF EAGLCLIENT}
           Q.TQ.Seek(FListe.Row-1) ;
         {$ENDIF}
         if Transactions(SetDuplic,0) <> oeOK then PGIBox('Impossible de dupliquer la pièce n° '+Q.FindField('GP_NUMERO').AsString, Caption);
         MoveCur(False);
       end;
       FListe.ClearSelected;
       FiniMove;
     end;
     ChercheClick;
   Finally
    bSelectAll.Down := False;
    SourisNormale;
    end;
   end;
   // Libération de la TOB Tiers
   if Tob_Tiers <> nil then Tob_Tiers.free;
   Tob_Tiers := nil ;
end ;

// Procédure appelée par une transaction pour exporter les lignes de la pièce
procedure TOF_MBORECALCULPIECE.SetDuplic;
var Nature,Souche,Nb,Indice : String;
    StWherePiece,StSQL : String;
    Q1 : TQuery;
    Tob_Piece : TOB ;
begin
  with TFMul(Ecran) do
  begin
    Nature  := Q.FindField('GP_NATUREPIECEG').AsString;
    Souche  := Q.FindField('GP_SOUCHE').AsString;
    Nb      := Q.FindField('GP_NUMERO').AsString;
    Indice  := Q.FindField('GP_INDICEG').AsString;
  end;
  StWherePiece := 'WHERE GP_NATUREPIECEG="'+Nature+'" AND '+
                  'GP_SOUCHE="'+Souche+'" AND GP_NUMERO='+Nb+' AND GP_INDICEG='+Indice;
  StSQL := 'SELECT * FROM PIECE '+ StWherePiece ;

  Q1 := OpenSQL(StSQL, True);
  if not Q1.Eof then
  begin
    // Chargement de la pièce
    Tob_Piece := TOB.Create ('PIECE', nil, -1);
    Tob_Piece.SelectDB('',Q1);
    // Duplication de la pièce
    DuplicPiece(Tob_Piece);
    // Libération de la pièce
    Tob_Piece.free ;
    Tob_Piece := nil ;
  end;
  Ferme(Q1);
end;

// Procédure appelée par une transactions pour recalculer toutes les pièces
procedure TOF_MBORECALCULPIECE.SetAllDuplic;
var StSQL, StSQL2 : string;
    Q1 : TQuery;
    Q2 : TQuery;
    Tob_Piece : TOB ;
    Nature,Souche,Nb,Indice, StWherePiece : String;
begin
  // Chargement de toutes les entêtes de pièce sélectionnées
  StSQL := 'SELECT GP_NATUREPIECEG, GP_SOUCHE, GP_NUMERO, GP_INDICEG FROM PIECE ' + StWhere ;

  Q1 := OpenSQL(StSQL, True);
  if not Q1.Eof then
  begin
     Q1.First;
     While Not Q1.Eof do
     begin
       // Chargement de la pièce
       Nature  := Q1.FindField('GP_NATUREPIECEG').AsString;
       Souche  := Q1.FindField('GP_SOUCHE').AsString;
       Nb      := Q1.FindField('GP_NUMERO').AsString;
       Indice  := Q1.FindField('GP_INDICEG').AsString;

       StWherePiece := 'WHERE GP_NATUREPIECEG="'+Nature+'" AND '+
                              'GP_SOUCHE="'+Souche+'" AND GP_NUMERO='+Nb+' AND GP_INDICEG='+Indice;

       StSQL2 := 'SELECT * FROM PIECE '+  StWherePiece ;
       Q2 := OpenSQL(StSQL2, True);

       if not Q2.Eof then
         begin
         // Chargement de la pièce
         Tob_Piece := TOB.Create ('PIECE', nil, -1);
         Tob_Piece.SelectDB('',Q2);
         //Duplication de la pièce
         DuplicPiece (Tob_Piece);
         // Libération de la pièce
         Tob_Piece.free; Tob_Piece := nil ;
         end;
       Ferme (Q2);
       // Pièce suivante
       Q1.Next;
     end;
  end;
  Ferme(Q1);
end;

function TOF_MBORECALCULPIECE.ExistePiece(Nature,Souche : string; Numero:integer): boolean;
begin
Result := ExisteSQL( 'Select GP_NATUREPIECEG from PIECE where GP_NATUREPIECEG="'+Nature+'" '+
                     'AND GP_SOUCHE="'+Souche+'" AND GP_NUMERO='+IntToStr(Numero));
end;

procedure TOF_MBORECALCULPIECE.DuplicPiece(var Tob_Piece : TOB);
var //Tob_Article     : TOB       ;
    Tob_Ligne,TobFilleArticle : TOB;
    Tob_Piedbase,Tob_PiedEche : TOB;
    Nature, Souche, UneLigne, SQL : string;
    Numero, indice, i  : integer;
    Condition,Depot,CodeArt : string;
    Q2              : Tquery    ;
    cpt             : integer   ;
begin
    /////////////////////////////////////////////////////////////////////////
    // Chargement de la TOB des lignes du document
    // (qui doivent être fille de l'entête)
    /////////////////////////////////////////////////////////////////////////
    Nature   := TOB_Piece.GetValue ('GP_NATUREPIECEG');
    Souche   := TOB_Piece.GetValue ('GP_SOUCHE');
    Numero   := TOB_Piece.GetValue ('GP_NUMERO');
    Indice   := TOB_Piece.GetValue ('GP_INDICEG');

    if Nature='TEM' then
      begin if ExistePiece('TRE',Souche,Numero) then exit; end
    else
      begin if ExistePiece('TEM',Souche,Numero) then exit; end;

    Condition:= 'GL_NATUREPIECEG="'+Nature+'" AND GL_SOUCHE="'+Souche+'" AND GL_NUMERO="'+IntToStr(Numero)+'" AND GL_INDICEG="'+IntToStr(Indice)+'"';
    SQL:='SELECT * From LIGNE WHERE ' + Condition ;
    Q2:=OpenSQL(SQL,True) ;
    if Not Q2.EOF then TOB_Piece.LoadDetailDB('LIGNE','','',Q2,True,True) ;
    Ferme (Q2) ;

    {
    CodeArt := '' ;
    // Création et chargement TOB TIERS
    CodeTiers := Tob_Piece.GetValue ('GP_TIERS');
    if CodeTiers <> CodeTiersOld then
      begin
      if Tob_Tiers <> nil then
        begin Tob_Tiers.free; Tob_Tiers := nil; end;
      SQL:='Select * From TIERS WHERE T_TIERS="'+CodeTiers+'"';
      Q2:=OpenSQL(SQL,True) ;
      if Not Q2.EOF then
        begin
        Tob_Tiers := TOB.Create ('TIERS', nil, -1);
        Tob_Tiers.SelectDB('',Q2);
        end;
      Ferme (Q2);
      // Sauvegarde du code tiers
      CodeTiersOld := CodeTiers ;
      end;
    /////////////////////////////////////////////////////////////////////////
    // Chargement de la TOB des articles + Fiches DISPO
    /////////////////////////////////////////////////////////////////////////
    Tob_Article := TOB.CREATE ('Les Articles', nil, -1);
    for Cpt:=0 to Tob_Piece.Detail.Count-1 do
    begin
      Tob_Ligne := Tob_Piece.Detail[Cpt];
      AddLesSupLigne(Tob_Ligne,False) ;

      CodeArt   := Tob_Ligne.GetValue ('GL_ARTICLE');
      if CodeArt <> '' then
      begin
        SQL:='Select * From ARTICLE WHERE GA_ARTICLE="'+CodeArt+'"';
        Q2:=OpenSQL(SQL,True) ;
        if Not Q2.EOF then
          begin
          TobFilleArticle := CreerTOBArt(Tob_Article);
          TobFilleArticle.SelectDB('',Q2);
          //LoadTOBDispo (TobFilleArticle, True, CreerQuelDepot(TobEntete)) ;
          end;
        Ferme (Q2);
      end;
    end;
    }

    /////////////////////////////////////////////////////////////////////////
    // Chargement de Piedbase
    /////////////////////////////////////////////////////////////////////////
    Tob_PiedBase := TOB.CREATE ('Les PiedsBases', nil, -1);
    Condition:= 'GPB_NATUREPIECEG="'+Nature+'" AND GPB_SOUCHE="'+Souche+'" AND GPB_NUMERO="'+IntToStr(Numero)+'" AND GPB_INDICEG="'+IntToStr(Indice)+'"';
    SQL:='SELECT * From PIEDBASE WHERE '+Condition ;
    Q2:=OpenSQL(SQL,True) ;
    if Not Q2.EOF then
      begin Tob_PiedBase.LoadDetailDB('PIEDBASE','','',Q2,True,True); end;
    Ferme(Q2) ;

    /////////////////////////////////////////////////////////////////////////
    // Chargement de Piedbase
    /////////////////////////////////////////////////////////////////////////
    Tob_PiedEche := TOB.CREATE ('Les PiedsEche', nil, -1);
    Condition:= 'GPE_NATUREPIECEG="'+Nature+'" AND GPE_SOUCHE="'+Souche+'" AND GPE_NUMERO="'+IntToStr(Numero)+'" AND GPE_INDICEG="'+IntToStr(Indice)+'"';
    SQL:='SELECT * From PIEDECHE WHERE '+Condition ;
    Q2:=OpenSQL(SQL,True) ;
    if Not Q2.EOF then
      begin Tob_PiedEche.LoadDetailDB('PIEDECHE','','',Q2,True,True); end;
    Ferme(Q2) ;

    //Transforme la pièce TEM en TRE et inversement
    if Nature='TEM' then
      begin
      Depot:=Tob_Piece.GetValue('GP_DEPOTDEST');
      Tob_Piece.PutValue('GP_NATUREPIECEG','TRE');
      end
    else
      begin
      Depot:=Tob_Piece.GetValue('GP_DEPOT');
      Tob_Piece.PutValue('GP_NATUREPIECEG','TEM');
      end;
    For i:=0 to Tob_Piece.Detail.Count-1 do
      begin
      Tob_Piece.Detail[i].PutValue('GL_NATUREPIECEG',Tob_Piece.GetValue('GP_NATUREPIECEG'));
      Tob_Piece.Detail[i].PutValue('GL_DEPOT',Depot);
      Tob_Piece.Detail[i].PutValue('GL_ETABLISSEMENT',Depot);
      end;
    For i:=0 to Tob_PiedBase.Detail.Count-1 do
      Tob_Piedbase.Detail[i].PutValue('GPB_NATUREPIECEG',Tob_Piece.GetValue('GP_NATUREPIECEG'));
    For i:=0 to Tob_PiedEche.Detail.Count-1 do
      Tob_PiedEche.Detail[i].PutValue('GPE_NATUREPIECEG',Tob_Piece.GetValue('GP_NATUREPIECEG'));

    // MAJ
    Tob_Piece.SetAllModifie(True) ;
    Tob_PiedBase.SetAllModifie(True) ;
    Tob_PiedEche.SetAllModifie(True) ;

    Tob_Piece.InsertDB(Nil);
    Tob_PiedBase.InsertDB(Nil);
    Tob_PiedEche.InsertDB(Nil);

    //if Tob_Article <> nil then begin Tob_Article.free ; Tob_Article := nil; end;
    if Tob_PiedBase <> nil then begin Tob_PiedBase.free; Tob_PiedBase:= nil; end;
    if Tob_PiedEche <> nil then begin Tob_PiedEche.free; Tob_PiedEche:= nil; end;
end;

procedure TOF_MBORECALCULPIECE.OnLoad ;
begin
  Inherited ;
  stWhere := RecupWhereCritere(TPageControl(TFMul(Ecran).Pages));
end ;

// procedure appellée par le bouton BOuvrir
procedure AGLDuplicPiece(Parms : Array of Variant; Nb : Integer);
var F : TForm;
    TOTOF : TOF;
begin
F := TForm(Longint(Parms[0]));
if (F is TFmul) then TOTOF := TFMul(F).LaTOF else exit;
if (TOTOF is TOF_MBORECALCULPIECE) then TOF_MBORECALCULPIECE(TOTOF).BatchDuplic;
end;

Initialization
  registerclasses ( [TOF_MBORECALCULPIECE] ) ;
  RegisterAGLProc('BatchRecalcul', True, 0, AGLDuplicPiece);
end.
