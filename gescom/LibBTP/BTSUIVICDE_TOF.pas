{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 21/06/2016
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSUIVICDE ()
Mots clefs ... : TOF;BTSUIVICDE
*****************************************************************}
Unit BTSUIVICDE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     uTob,
{$ENDIF}
     forms,
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1, 
     HMsgBox,
     UTOB,
     HPanel,
     UTOF,
     uEntCommun,
     UCotraitance,
     Vierge,
     HSysMenu,
     HTB97,
     UtilsGrille;


Type
  TOF_BTSUIVICDE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private

    //
    PGRILLE       : TPageControl;
    //
    TabLigCde     : TTabSheet;
    TabReception  : TTabSheet;
    TabFacture    : TTabSheet;
    //
    Panel1        : THPanel;
    Panel2        : THPanel;
    Panel3        : THPanel;
    Panel4        : THPanel;
    Panel5        : THPanel;
    //
    Numero        : THEdit;
    DateCde       : THEdit;
    Fournisseur   : THEdit;
    Chantier      : THEdit;
    Chantier0     : THEdit;
    Chantier1     : THEdit;
    Chantier2     : THEdit;
    Chantier3     : THEdit;
    Chantier4     : THEdit;
    REFInterne    : THEdit;
    DateLiv       : THEdit;
    Etablissement : THEdit;
    Devise        : THEdit;
    contact       : THEdit;
    Depot         : THEdit;
    Domaine       : THEdit;
    TotalHT       : THEdit;
    TotalTVA      : THEdit;
    TotalTTC      : THEdit;
    //
    TNomFrs       : THLabel;
    TAdr1Frs      : THLabel;
    TAdr2Frs      : THLabel;
    TAdr3Frs      : THLabel;
    TCPFrs        : THLabel;
    TVilleFrs     : THLabel;
    TLibChantier  : THLabel;
    TLibEtab      : THLabel;
    TLibDevise    : THLabel;
    TLibContact   : THLabel;
    TLibDepot     : THLabel;
    TLibDomaine   : THLabel;
    //
    GrilleLigCde  : THGrid;
    GrilleEntRec  : THGrid;
    GrilleLigRec  : THGrid;
    GrilleEntFac  : THGrid;
    GrilleLigFac  : THGrid;
    //
    GGrilleLigCde  : TGestionGS;
    GGrilleEntRec  : TGestionGS;
    GGrilleLigRec  : TGestionGS;
    GGrilleEntFac  : TGestionGS;
    GGrilleLigFac  : TGestionGS;
    //
    BtParamList     : TToolbarButton97;
    //
    TobEntCde       : Tob;
    TobLigCde       : Tob;
    TobEntRec       : Tob;
    TobLigRec       : Tob;
    TobEntFac       : Tob;
    TobLigFac       : Tob;
    //
    Cledoc          : R_CLEDOC;
    //
    PieceOrigine    : String;
    PiecePrecedente : String;
    StSQL           : String;
    QQ              : TQuery;
    //
    procedure BtParamListOnClick(Sender: Tobject);
    procedure CalculMontantCde;
    procedure CalculMontantFac;
    procedure CalculMontantRec;
    procedure ChargementEcran;
    procedure ChargementLigneFacturation;
    procedure ChargementLigneReception;
    procedure ChargementTOB;
    procedure ClickGrilleEntFac(Sender: TObject);
    procedure ClickGrilleEntRec(Sender: TObject);
    procedure Controlechamp(Champ, Valeur: String);
    procedure CreateTOB;
    procedure DblClickGrilleEntFac(Sender: TObject);
    procedure DblClickGrilleEntRec(Sender: TObject);
    procedure DestroyTOB;
    procedure GetObjects;
    procedure InitGrilleEntFac;
    procedure InitGrilleEntRec;
    procedure InitGrilleLigCde;
    procedure InitGrilleLigFac;
    procedure InitGrilleLigRec;
    procedure PGrilleOnChange(Sender: TObject);
    procedure SetScreenEvents;
    //
  end ;

Implementation
Uses  affaireutil,
      FactComm,
      Facture,
      StrUtils,
      AGLInitGC,
      ParamDBG,
      TntGrids, Grids;

procedure TOF_BTSUIVICDE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSUIVICDE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSUIVICDE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTSUIVICDE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTSUIVICDE.OnArgument (S : String ) ;
var x       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
begin
  Inherited ;
  //
  //Chargement des zones ecran dans des zones programme
  GetObjects;
  //
  CreateTOB;
  //
  Critere := uppercase(Trim(ReadTokenSt(S)));
  while Critere <> '' do
  begin
     x := pos('=', Critere);
     if x <> 0 then
        begin
        Champ  := copy(Critere, 1, x - 1);
        Valeur := copy(Critere, x + 1, length(Critere));
        end
     else
        Champ := Critere;
     ControleChamp(Champ, Valeur);
     Critere:= uppercase(Trim(ReadTokenSt(S)));
  end;

  //Gestion des évènement des zones écran
  SetScreenEvents;

  //Initialisation des grilles
  InitGrilleLigCde;
  InitGrilleEntRec;
  InitGrilleLigRec;
  InitGrilleEntFac;
  InitGrilleLigFac;

  //chargement des tobs
  ChargementTOB;

  //chargement des zones écran avec les zones tob
  ChargementEcran;

  //Colonne proportionnelles à l'écran
  THSystemMenu(getControl('HMtrad')).ResizeGridColumns(GrilleLigCde);
  THSystemMenu(getControl('HMtrad')).ResizeGridColumns(GrilleLigRec);
  THSystemMenu(getControl('HMtrad')).ResizeGridColumns(GrilleLigFac);
  //
  PGRILLE.ActivePageIndex := 0;
  //
  CalculMontantCde;

end ;

procedure TOF_BTSUIVICDE.GetObjects ;
begin
    //
    PGRILLE       := TPageControl(GetControl('PGRILLE'));
    //
    TabLigCde     := TTabSheet(GetControl('TABLIGCDE'));
    TabReception  := TTabSheet(GetControl('TABRECEPTION'));
    TabFacture    := TTabSheet(GetControl('TABFACTURE'));
    //
    Panel1        := THPanel(GetControl('PANEL1'));
    Panel2        := THPanel(GetControl('PANEL2'));
    Panel3        := THPanel(GetControl('PANEL3'));
    Panel4        := THPanel(GetControl('PANEL4'));
    Panel5        := THPanel(GetControl('PANEL5'));
    //
    Numero        := THEdit(GetControl('NUMERO'));
    DateCde       := THEdit(GetControl('DATECDE'));
    Fournisseur   := THEdit(GetControl('FOURNISSEUR'));
    Chantier      := THEdit(GetControl('CHANTIER'));
    Chantier0     := THEdit(GetControl('CHANTIER0'));
    Chantier1     := THEdit(GetControl('CHANTIER1'));
    Chantier2     := THEdit(GetControl('CHANTIER2'));
    Chantier3     := THEdit(GetControl('CHANTIER3'));
    Chantier4     := THEdit(GetControl('AVENANT'));
    REFInterne    := THEdit(GetControl('REFINTERNE'));
    DateLiv       := THEdit(GetControl('DATELIV'));
    Etablissement := THEdit(GetControl('ETABLISSEMENT'));
    Devise        := THEdit(GetControl('DEVISe'));
    contact       := THEdit(GetControl('CONTACT'));
    Depot         := THEdit(GetControl('DEPOT'));
    Domaine       := THEdit(GetControl('DOMAINE'));
    //
    TotalHT       := THEdit(GetControl('TOTALHT'));
    TotalTVA      := THEdit(GetControl('TOTALTVA'));
    TotalTTC      := THEdit(GetControl('TOTALTTC'));
    //
    TNomFrs       := THLabel(GetControl('TNOMFRS'));
    TAdr1Frs      := THLabel(GetControl('TADR1FRS'));
    TAdr2Frs      := THLabel(GetControl('TADR2FRS'));
    TAdr3Frs      := THLabel(GetControl('TADR3FRS'));
    TCPFrs        := THLabel(GetControl('TCPFRS'));
    TVilleFrs     := THLabel(GetControl('TVILLEFRS'));
    TLibChantier  := THLabel(GetControl('TLIBAFFAIRE'));
    TLibEtab      := THLabel(GetControl('TLIBETABLISSEMENT'));
    TLibDevise    := THLabel(GetControl('TLIBDEVISE'));
    TLibContact   := THLabel(GetControl('TLIBCONTACT'));
    TLibDepot     := THLabel(GetControl('TLIBDEPOT'));
    TLibDomaine   := THLabel(GetControl('TLIBDOMAINE'));
    //
    GrilleLigCde  := THGrid(GetControl('GRILLELIGCDE'));
    GrilleEntRec  := THGrid(GetControl('GRILLEENTRECEP'));
    GrilleLigRec  := THGrid(GetControl('GRILLELIGRECEP'));
    GrilleEntFac  := THGrid(GetControl('GRILLEENTFACT'));
    GrilleLigFac  := THGrid(GetControl('GRILLELIGFACT'));
    //
    BtParamList   := TToolbarButton97(GetControl('BPARAMLISTE'));

end;

Procedure TOF_BTSUIVICDE.CreateTOB ;
begin

    TobEntCde     := Tob.Create('ENTCDE', nil, -1);
    TobLigCde     := Tob.Create('LIGCDE', nil, -1);
    TobEntRec     := Tob.Create('ENTREC', nil, -1);
    TobLigRec     := Tob.Create('LIGREC', nil, -1);
    TobEntFac     := Tob.Create('ENTFAC', nil, -1);
    TobLigFac     := Tob.Create('LIGFAC', nil, -1);

end;

Procedure TOF_BTSUIVICDE.DestroyTOB ;
begin

    FreeAndNil(TobEntCde);
    FreeAndNil(TobLigCde);
    FreeAndNil(TobEntRec);
    FreeAndNil(TobLigRec);
    FreeAndNil(TobEntFac);
    FreeAndNil(TobLigFac);

    //Destruction des objets... c'était plus simple de les mettre ici !!!
    FreeAndNil(GGrilleLigCde);
    FreeAndNil(GGrilleEntRec);
    FreeAndNil(GGrilleLigRec);
    FreeAndNil(GGrilleEntFac);
    FreeAndNil(GGrilleLigFac);

end;

Procedure TOF_BTSUIVICDE.ChargementTOB ;
Var NumReception  : Integer;
    NumFacture    : Integer;
    TobLentRec    : TOB;
    Q1            : TQuery;
begin

  //chargement de la tob Entête de commande
  StSQL := 'SELECT AFF_LIBELLE, ';
  StSQL := StSQL  + ' T_LIBELLE,T_ADRESSE1,T_ADRESSE2,T_ADRESSE3, T_CODEPOSTAL, T_VILLE, ';
  StSQL := StSQL  + ' ARS_LIBELLE, ET_LIBELLE, D_LIBELLE, GDE_LIBELLE, BTD_LIBELLE, PIECE.* ';
  StSQL := StSQL  + ' FROM PIECE ';
  StSQL := StSQL  + ' LEFT JOIN AFFAIRE       ON GP_AFFAIRE=AFF_AFFAIRE ';
  StSQL := StSQL  + ' LEFT JOIN TIERS         ON GP_TIERS=T_TIERS ';
  StSQL := StSQL  + ' LEFT JOIN RESSOURCE     ON GP_RESSOURCE=ARS_RESSOURCE';
  StSQL := StSQL  + ' LEFT JOIN ETABLISS      ON GP_ETABLISSEMENT=ET_ETABLISSEMENT ';
  StSQL := StSQL  + ' LEFT JOIN DEVISE        ON GP_DEVISE=D_DEVISE ';
  StSQL := StSQL  + ' LEFT JOIN DEPOTS        ON GP_DEPOT=GDE_DEPOT';
  StSQL := StSQL  + ' LEFT JOIN BTDOMAINEACT  ON GP_DOMAINE=BTD_CODE ';
  StSQL := StSQL  + 'WHERE GP_NATUREPIECEG="' + Cledoc.NaturePiece;
  StSQL := StSQL  + '" AND GP_SOUCHE="' + Cledoc.Souche;
  StSQL := StSQL  + '" AND GP_NUMERO='  + IntToStr(Cledoc.NumeroPiece);
  StSQL := StSQL  + '  AND GP_INDICEG=' + IntToStr(Cledoc.Indice);

  QQ := OpenSQL(StSQL, False, -1, '',True);
  IF not QQ.Eof then
  begin
    TobEntCde.SelectDB('PIECE',QQ,False);
  end;

  Cledoc.DatePiece := TobEntCde.GetDateTime('GP_DATEPIECE');
  //
  //PieceOrigine := EncodeRefPiece(Cledoc);
  PieceOrigine := FormatDateTime('ddmmyyyy', cledoc.DatePiece);
  PieceOrigine := PieceOrigine + ';' + cledoc.NaturePiece  + ';' + Cledoc.Souche  + ';' +  IntToStr(cledoc.NumeroPiece);

  Ferme(QQ);

  //chargement de la tob Lignes de commande
  StSQL := 'SELECT LIGNE.* ';
  StSQL := StSQL  + ' FROM LIGNE ';
  StSQL := StSQL  + 'WHERE GL_NATUREPIECEG="' + Cledoc.NaturePiece;
  StSQL := StSQL  + '" AND GL_SOUCHE="' + Cledoc.Souche;
  StSQL := StSQL  + '" AND GL_NUMERO='  + IntToStr(Cledoc.NumeroPiece);
  StSQL := StSQL  + '  AND GL_INDICEG=' + IntToStr(Cledoc.Indice);
  //StSQL := StSQL  + '  AND GL_TYPELIGNE="ART"';

  QQ := OpenSQL(StSQL, False, -1, '',True);
  IF not QQ.Eof then
  begin
    TobLigCde.LoadDetailDB('LIGNE','','',QQ,False);
  end;

  Ferme(QQ);

  //Recherche du N° de Réception...
  Try
    StSQL := 'SELECT DISTINCT GL_NUMERO FROM LIGNE ';
    StSQL := StSQL  + 'WHERE GL_PIECEPRECEDENTE LIKE "' + PieceOrigine + '%"';
    StSQL := StSQL  + '  AND GL_NATUREPIECEG="BLF"';
    QQ := OpenSQL(STSQL, False, -1, '', False);
    QQ.First;
    while Not QQ.EOF do
    Begin
      NumReception := QQ.Findfield('GL_NUMERO').AsInteger;

      //chargement de la Tob Entete de Réception
      StSQL := 'SELECT  ' + Copy(AnsiReplaceStr(GGrilleEntRec.ColNamesGS, ';', ','),0, Length(GGrilleEntRec.ColNamesGS)-2);
      StSQL := StSQL  + ' FROM BTDEVIS ';
      StSQL := StSQL  + 'WHERE GP_NUMERO = ' + IntToStr(NumReception) + '  AND GP_NATUREPIECEG="BLF"';
      Q1 := OpenSQL(StSQL, False, -1, '',False);
      IF not Q1.Eof then
      begin
        TobEntRec.LoadDetailDB('LES RECEPTIONS','','',Q1,True, True);
      end;
      Ferme(Q1);
      QQ.Next;
    end;
  finally
    Ferme(QQ);
  end;

  //chargement de la Tob Ligne de Réception
  if TobEntRec.Detail.count <> 0 then
  begin
    GrilleEntRec.Row := 1;
    ChargementLigneReception;
  end;

  Try
    //chargement de la Tob Entete de Facturation
    StSQL := 'SELECT DISTINCT GL_NUMERO FROM LIGNE ';
    StSQL := StSQL  + 'WHERE GL_PIECEORIGINE LIKE "' + PieceOrigine + '%"';
    StSQL := StSQL  + '  AND GL_NATUREPIECEG="FF"';
    QQ := OpenSQL(STSQL, False, -1, '', True);
    QQ.First;
    while Not QQ.EOF do
    Begin
      NumFacture := QQ.Findfield('GL_NUMERO').AsInteger;

      //chargement de la Tob Entete de Facturation
      StSQL := 'SELECT ' + Copy(AnsiReplaceStr(GGrilleEntfac.ColNamesGS, ';', ','),0, Length(GGrilleEntfac.ColNamesGS)-2);
      StSQL := StSQL  + ' FROM BTDEVIS ';
      StSQL := StSQL  + 'WHERE GP_NUMERO = ' + IntToStr(NumFacture) + ' AND GP_NATUREPIECEG="FF"';
      Q1 := OpenSQL(StSQL, False, -1, '',False);
      IF not Q1.Eof then
      begin
        TobEntfac.LoadDetailDB('LES FACTURES','','',Q1,True, True);
      end;
      Ferme(Q1);
      QQ.Next;
    end;
  finally
    Ferme(QQ);
  end;

  //chargement de la Tob Ligne de Facturation
  //FV1 - 03/11/2017 - FS#2763 - TESTS BL : 
  if TobEntfac.Detail.count <>0 then
  begin
    GrilleEntFac.Row := 1;
    ChargementLigneFacturation;
  end;

end;

Procedure TOF_BTSUIVICDE.ChargementLigneReception;
Var Reception : Tob;
begin

  if TobEntRec.detail.count = 0 then exit;

  PiecePrecedente  := FormatDateTime('ddmmyyyy',GGrilleEntRec.TOBG.detail[GrilleEntRec.row-1].GetDateTime('GP_DATEPIECE'));
  PiecePrecedente  := PiecePrecedente + ';' +   GGrilleEntRec.TOBG.detail[GrilleEntRec.row-1].GetString('GP_NATUREPIECEG');
  PiecePrecedente  := PiecePrecedente + ';' +   GGrilleEntRec.TOBG.detail[GrilleEntRec.row-1].GetString('GP_SOUCHE');
  PiecePrecedente  := PiecePrecedente + ';' +   GGrilleEntRec.TOBG.detail[GrilleEntRec.row-1].GetString('GP_NUMERO');

  //lecture des lignes associées à la sélection
  Reception := GGrilleEntRec.TOBG.detail[GrilleEntRec.row-1];

  if Reception = nil then Exit;

  //chargement de la tob Ligne de Réception en fonction de l'entête sélectionnée.
  //chargement de la tob Lignes de commande
  StSQL := 'SELECT LIGNE.* ';
  StSQL := StSQL  + ' FROM LIGNE ';
  StSQL := StSQL  + 'WHERE GL_NATUREPIECEG="' + Reception.GetValue('GP_NATUREPIECEG');
  StSQL := StSQL  + '" AND GL_SOUCHE="' + Reception.GetValue('GP_SOUCHE');
  StSQL := StSQL  + '" AND GL_NUMERO='  + IntToStr(Reception.GetValue('GP_NUMERO'));
  StSQL := StSQL  + '  AND GL_INDICEG=' + IntToStr(Reception.GetValue('GP_INDICEG'));
  //FV1 - 03/11/2017 - FS#2763 - TESTS BL :
  //StSQL := StSQL  + '  AND GL_PIECEORIGINE LIKE "' + PieceOrigine + '%"';
  StSQL := StSQL  + '  AND GL_TYPELIGNE="ART"';

  QQ := OpenSQL(StSQL, False, -1, '',True);
  IF not QQ.Eof then
  begin
    TobLigRec.LoadDetailDB('LIGNE RECEPTION','','',QQ,False);
  end;

  Ferme(QQ);

end;

Procedure TOF_BTSUIVICDE.ChargementLigneFacturation;
Var Facturation : Tob;
begin

  if TobEntFac.detail.count = 0 then exit;

  //lecture des lignes associées à la sélection
  Facturation := GGrilleEntFac.TOBG.detail[GrilleEntFac.row-1];

  if Facturation = nil then Exit;

  //chargement de la tob Ligne de Réception en fonction de l'entête sélectionnée.
  //chargement de la tob Lignes de commande
  StSQL := 'SELECT LIGNE.* ';
  StSQL := StSQL  + ' FROM LIGNE ';
  StSQL := StSQL  + 'WHERE GL_NATUREPIECEG="' + Facturation.GetValue('GP_NATUREPIECEG');
  StSQL := StSQL  + '" AND GL_SOUCHE="' + Facturation.GetValue('GP_SOUCHE');
  StSQL := StSQL  + '" AND GL_NUMERO='  + IntToStr(Facturation.GetValue('GP_NUMERO'));
  StSQL := StSQL  + '  AND GL_INDICEG=' + IntToStr(Facturation.GetValue('GP_INDICEG'));
  //FV1 - 03/11/2017 - FS#2763 - TESTS BL :
  //StSQL := StSQL  + '  AND GL_PIECEORIGINE LIKE "' + PieceOrigine + '%"';
  StSQL := StSQL  + '  AND GL_TYPELIGNE="ART"';

  QQ := OpenSQL(StSQL, False, -1, '',True);
  IF not QQ.Eof then
  begin
    TobLigFac.LoadDetailDB('LIGNE FACTURATION','','',QQ,False);
  end;

  Ferme(QQ);

end;

Procedure TOF_BTSUIVICDE.ChargementEcran;
begin

//Chargement des zones entête
  Numero.text        := TobEntCde.GetValue('GP_NUMERO');
  DateCde.text       := TobEntCde.GetValue('GP_DATEPIECE');
  Fournisseur.Text   := TobEntCde.GetValue('GP_TIERS');
  Chantier.Text      := TobEntCde.GetValue('GP_AFFAIRE');
  Chantier0.Text     := TobEntCde.GetValue('GP_AFFAIRE0');
  Chantier1.Text     := TobEntCde.GetValue('GP_AFFAIRE1');
  Chantier2.Text     := TobEntCde.GetValue('GP_AFFAIRE2');
  Chantier3.Text     := TobEntCde.GetValue('GP_AFFAIRE3');
  Chantier4.Text     := TobEntCde.GetValue('GP_AVENANT');

  if Chantier1.text = '' then
  begin
    Chantier.Visible  := False;
    Chantier0.Visible := False;
    Chantier1.Visible := False;
    Chantier2.Visible := False;
    Chantier3.Visible := False;
    Chantier4.Visible := False;
  end;

  REFInterne.Text       := TobEntCde.GetValue('GP_REFINTERNE');
  DateLiv.Text          := TobEntCde.GetValue('GP_DATELIVRAISON');
  Etablissement.Text    := TobEntCde.GetValue('GP_ETABLISSEMENT');
  Devise.Text           := TobEntCde.GetValue('GP_DEVISE');
  contact.Text          := TobEntCde.GetValue('GP_RESSOURCE');
  Depot.Text            := TobEntCde.GetValue('GP_DEPOT');
  Domaine.Text          := TobEntCde.GetValue('GP_DOMAINE');
  //
  TNomFrs.caption       := TobEntCde.GetValue('T_LIBELLE');
  TAdr1Frs.caption      := TobEntCde.GetValue('T_ADRESSE1');
  TAdr2Frs.caption      := TobEntCde.GetValue('T_ADRESSE2');
  TAdr3Frs.caption      := TobEntCde.GetValue('T_ADRESSE3');
  TCPFrs.caption        := TobEntCde.GetValue('T_CODEPOSTAL');
  TVilleFrs.caption     := TobEntCde.GetValue('T_VILLE');
  TLibChantier.caption  := TobEntCde.GetValue('AFF_LIBELLE');
  TLibEtab.caption      := TobEntCde.GetValue('ET_LIBELLE');
  TLibDevise.caption    := TobEntCde.GetValue('D_LIBELLE');
  TLibContact.caption   := TobEntCde.GetValue('ARS_LIBELLE');
  TLibDepot.caption     := TobEntCde.GetValue('GDE_LIBELLE');
  TLibDomaine.caption   := TobEntCde.GetValue('BTD_LIBELLE');
  //
  ChargeCleAffaire(CHANTIER0,CHANTIER1,CHANTIER2,CHANTIER3,CHANTIER4,Nil,taconsult,CHANTIER.Text,False);
  //
  if TobLigCde.Detail.count <>0 then
  begin
    GGrilleLigCde.ChargementGrille;
  end;

  if TobEntRec.Detail.count <>0 then
  begin
    GGrilleEntRec.ChargementGrille;
    GrilleEntRec.Row := 1;
    GGrilleLigRec.ChargementGrille;
  end
  else
    TabReception.TabVisible := False;

  if TobEntFac.Detail.count <>0 then
  begin
    GGrilleEntFac.ChargementGrille;
    GrilleEntFac.Row := 1;
    GGrilleLigFac.ChargementGrille;
  end
  else
    TabFacture.TabVisible := False;

end;

procedure TOF_BTSUIVICDE.ClickGrilleEntRec(Sender : TObject);
begin

  ChargementLigneReception;

  if TobLigRec.Detail.count > 0 then
  begin
    GGrilleLigRec.ChargementGrille;
  end;

  ChargementLigneFacturation;

  CalculMontantRec;

end;

procedure TOF_BTSUIVICDE.ClickGrilleEntFac(Sender : TObject);
begin

  ChargementLigneFacturation;

  if TobLigFac.Detail.count >0 then
  begin
    GGrilleLigFac.ChargementGrille;
  end;

  CalculMontantFac;

end;

procedure TOF_BTSUIVICDE.DblClickGrilleEntRec(Sender : TObject);
Var ClePiece			: array [0..7] of Variant;
    TheChaine     : string;
begin

  ClePiece[1] := TobEntRec.Detail[GrilleEntRec.Row-1].GetValue('GP_NATUREPIECEG');
  ClePiece[2] := DateTimetoStr(TobEntRec.Detail[GrilleEntRec.Row-1].GetValue('GP_DATEPIECE'));
  ClePiece[3] := TobEntRec.Detail[GrilleEntRec.Row-1].GetValue('GP_SOUCHE');
  ClePiece[4] := IntToStr(TobEntRec.Detail[GrilleEntRec.Row-1].GetValue('GP_NUMERO'));
  ClePiece[5] := IntToStr(TobEntRec.Detail[GrilleEntRec.Row-1].GetValue('GP_INDICEG'));
  ClePiece[6] := TobEntRec.Detail[GrilleEntRec.Row-1].GetValue('GP_TIERS');

  if TobEntRec.Detail[GrilleEntRec.Row-1].GetValue('GP_AFFAIRE') <> '' then
  begin
    ClePiece[7] := TobEntRec.Detail[GrilleEntRec.Row-1].GetValue('GP_AFFAIRE');
    TheChaine := ClePiece[1]+';'+ClePiece[2]+';'+ClePiece[3]+';'+ClePiece[4]+';'+ClePiece[5]+';';
    AppelPiece([TheChaine,'ACTION=CONSULTATION',False],3);
  end
  else
  begin
    TheChaine := ClePiece[1]+';'+ClePiece[2]+';'+ClePiece[3]+';'+ClePiece[4]+';'+ClePiece[5]+';';
    AppelPiece([TheChaine,'ACTION=CONSULTATION',True],3);
  end;

end;

procedure TOF_BTSUIVICDE.DblClickGrilleEntFac(Sender : TObject);
Var ClePiece			: array [0..7] of Variant;
    TheChaine     : string;
begin

  ClePiece[1] := TobEntFac.Detail[GrilleEntFac.Row-1].GetValue('GP_NATUREPIECEG');
  ClePiece[2] := DateTimetoStr(TobEntFac.Detail[GrilleEntFac.Row-1].GetValue('GP_DATEPIECE'));
  ClePiece[3] := TobEntFac.Detail[GrilleEntFac.Row-1].GetValue('GP_SOUCHE');
  ClePiece[4] := IntToStr(TobEntFac.Detail[GrilleEntFac.Row-1].GetValue('GP_NUMERO'));
  ClePiece[5] := IntToStr(TobEntFac.Detail[GrilleEntFac.Row-1].GetValue('GP_INDICEG'));
  ClePiece[6] := TobEntFac.Detail[GrilleEntFac.Row-1].GetValue('GP_TIERS');

  if TobEntFac.Detail[GrilleEntFac.Row-1].GetValue('GP_AFFAIRE') <> '' then
  begin
    ClePiece[7] := TobEntFac.Detail[GrilleEntFac.Row-1].GetValue('GP_AFFAIRE');
    TheChaine := ClePiece[1]+';'+ClePiece[2]+';'+ClePiece[3]+';'+ClePiece[4]+';'+ClePiece[5]+';';
    AppelPiece([TheChaine,'ACTION=CONSULTATION',False],3);
  end
  else
  begin
    TheChaine := ClePiece[1]+';'+ClePiece[2]+';'+ClePiece[3]+';'+ClePiece[4]+';'+ClePiece[5]+';';
    AppelPiece([TheChaine,'ACTION=CONSULTATION',True],3);
  end;


end;

Procedure TOF_BTSUIVICDE.SetScreenEvents;
begin

  PGRILLE.OnChange           := PGrilleOnChange;

  GrilleEntRec.OnClick       := ClickGrilleEntRec;
  GrilleEntFac.OnClick       := ClickGrilleEntFac;

  GrilleEntRec.OnDblClick    := DblClickGrilleEntRec;
  GrilleEntFac.OnDblClick    := DblClickGrilleEntFac;

  BtParamList.OnClick        := BtParamListOnClick;

end;

Procedure TOF_BTSUIVICDE.Controlechamp(Champ, Valeur : String);
begin

  if Champ='DOCUMENT' then DecodeCleDoc(Valeur, Cledoc);

end;

procedure TOF_BTSUIVICDE.OnClose ;
begin
  Inherited ;

  DestroyTOB;
  
end ;

procedure TOF_BTSUIVICDE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSUIVICDE.OnCancel () ;
begin
  Inherited ;
end ;

Procedure TOF_BTSUIVICDE.InitGrilleLigCde;
begin

  //Une recherche de la grille au niveau de la table des liste serait bien venu !!!
  GGrilleLigCde           := TGestionGS.Create;

  GGrilleLigCde.Ecran     := TFVierge(Ecran);
  GGrilleLigCde.GS        := GrilleLigCde;
  GGrilleLigCde.TOBG      := TobLigCde;

  GGrilleLigCde.NomListe  := 'BTSAISIECF';

  GGrilleLigCde.ChargeInfoListe;

  GGrilleLigCde.RowHeight := 18;

  if GGrilleLigCde.NomListe <> '' then GGrilleLigCde.DessineGrille;

end;

Procedure TOF_BTSUIVICDE.InitGrilleEntRec;
begin

  //Une recherche de la grille au niveau de la table des liste serait bien venu !!!
  GGrilleEntRec           := TGestionGS.Create;

  GGrilleEntRec.Ecran     := TFVierge(Ecran);
  GGrilleEntRec.GS        := GrilleEntRec;
  GGrilleEntRec.TOBG      := TobEntRec;

  GGrilleEntRec.NomListe  := 'BTMULPIECE';

  GGrilleEntRec.ChargeInfoListe;

  GGrilleEntrec.RowHeight := 18;

  if GGrilleEntrec.NomListe <> '' then GGrilleEntRec.DessineGrille;

  THSystemMenu(getControl('HMtrad')).ResizeGridColumns(GrilleEntRec);
                
end;

Procedure TOF_BTSUIVICDE.InitGrilleLigRec;
begin

  //Une recherche de la grille au niveau de la table des liste serait bien venu !!!
  GGrilleLigRec           := TGestionGS.Create;

  GGrilleLigRec.Ecran     := TFVierge(Ecran);
  GGrilleLigRec.GS        := GrilleLigRec;
  GGrilleLigRec.TOBG      := TobLigRec;

  GGrilleLigRec.NomListe  := 'BTSAISIEBLF';

  GGrilleLigRec.ChargeInfoListe;

  GGrilleLigrec.RowHeight := 18;

  if GGrilleLigrec.NomListe <> '' then GGrilleLigRec.DessineGrille;

end;

Procedure TOF_BTSUIVICDE.InitGrilleEntFac;
begin

  //Une recherche de la grille au niveau de la table des liste serait bien venu !!!
  GGrilleEntFac           := TGestionGS.Create;

  GGrilleEntFac.Ecran     := TFVierge(Ecran);
  GGrilleEntFac.GS        := GrilleEntFac;
  GGrilleEntFac.TOBG      := TobEntFac;

  GGrilleEntFac.NomListe  := 'BTMULPIECE';

  GGrilleEntFac.ChargeInfoListe;

  GGrilleEntFac.RowHeight := 18;

  if GGrilleEntFac.NomListe <> '' then GGrilleEntFac.DessineGrille;

  THSystemMenu(getControl('HMtrad')).ResizeGridColumns(GrilleEntFac);

end;

Procedure TOF_BTSUIVICDE.InitGrilleLigFac;
begin

  //Une recherche de la grille au niveau de la table des liste serait bien venu !!!
  GGrilleLigFac           := TGestionGS.Create;

  GGrilleLigFac.Ecran     := TFVierge(Ecran);
  GGrilleLigFac.GS        := GrilleLigFac;
  GGrilleLigFac.TOBG      := TobLigFac;

  GGrilleLigFac.NomListe  := 'BTSAISIEFAF';

  GGrilleLigFac.ChargeInfoListe;

  GGrilleLigFac.RowHeight := 18;

  if GGrilleLigFac.NomListe <> '' then GGrilleLigFac.DessineGrille;

end;

procedure TOF_BTSUIVICDE.PGrilleOnChange(Sender : TObject);
begin

  IF PGRILLE.ActivePageIndex = 0 then
  begin
    CalculMontantCde;
    BtParamList.Visible := True;
  end
  else if PGRILLE.ActivePageIndex = 1 then
  Begin
    CalculMontantRec;
    BtParamList.Visible := True;
  end
  else if PGRILLE.ActivePageIndex = 2 then
  begin
    CalculMontantFac;
    BtParamList.Visible := True;
  end;

end;

Procedure TOF_BTSUIVICDE.CalculMontantCde;
Var MontantHT   : Double;
    MontantTVA  : Double;
    MontantTTC  : Double;
    Indice      : Integer;
Begin

  MontantHT   := 0;
  MontantTVA  := 0;
  MontantTTC  := 0;

  for Indice  := 0 to TobLigCde.detail.count -1 do
  begin
    MontantHT   := MontantHT  + TobLigCde.detail[Indice].GetDouble('GL_MONTANTHTDEV');
    MontantTTC  := MontantTTC + TobLigCde.detail[Indice].GetDouble('GL_MONTANTTTCDEV');
    MontantTVA  := MontantTTC - MontantHT;
  end;

  TotalHT.Text  := FloatToStr(MontantHT);
  TotalTVA.Text := FloatToStr(MontantTVA);
  TotalTTC.Text := FloatToStr(MontantTTC);

end;

Procedure TOF_BTSUIVICDE.CalculMontantRec;
Var MontantHT   : Double;
    MontantTVA  : Double;
    MontantTTC  : Double;
    Indice      : Integer;
Begin

  MontantHT   := 0;
  MontantTVA  := 0;
  MontantTTC  := 0;

  for Indice  := 0 to TobLigRec.detail.count -1 do
  begin
    MontantHT   := MontantHT  + TobLigRec.detail[Indice].GetDouble('GL_MONTANTHTDEV');
    MontantTTC  := MontantTTC + TobLigRec.detail[Indice].GetDouble('GL_MONTANTTTCDEV');
    MontantTVA  := MontantTTC - MontantHT;
  end;

  TotalHT.Text  := FloatToStr(MontantHT);
  TotalTVA.Text := FloatToStr(MontantTVA);
  TotalTTC.Text := FloatToStr(MontantTTC);

end;

Procedure TOF_BTSUIVICDE.CalculMontantFac;
Var MontantHT   : Double;
    MontantTVA  : Double;
    MontantTTC  : Double;
    Indice      : Integer;
Begin

  MontantHT   := 0;
  MontantTVA  := 0;
  MontantTTC  := 0;

  for Indice  := 0 to TobLigFac.detail.count -1 do
  begin
    MontantHT   := MontantHT  + TobLigFac.detail[Indice].GetDouble('GL_MONTANTHTDEV');
    MontantTTC  := MontantTTC + TobLigFac.detail[Indice].GetDouble('GL_MONTANTTTCDEV');
    MontantTVA  := MontantTTC - MontantHT;
  end;

  TotalHT.Text  := FloatToStr(MontantHT);
  TotalTVA.Text := FloatToStr(MontantTVA);
  TotalTTC.Text := FloatToStr(MontantTTC);

end;

procedure TOF_BTSUIVICDE.BtParamListOnClick(Sender: Tobject);
begin

  if PGrille.ActivePageIndex = 0 then
  begin
    ParamListe(GGrilleLigCde.NomListe, nil, nil, '');
    //on doit recharger les grilles !!!
    FreeAndNil(GGrilleLigCde);
    InitGrilleLigCde;
    GGrilleLigCde.ChargementGrille;
  end
  else if PGrille.ActivePageIndex = 1 then
  begin
    ParamListe(GGrilleLigRec.NomListe, nil, nil, '');
    //on doit recharger les grilles !!!
    FreeAndNil(GGrilleLigRec);
    InitGrilleLigRec;
    GGrilleLigRec.ChargementGrille;
  end
  else if PGrille.ActivePageIndex = 2 then
  begin
    ParamListe(GGrilleLigFac.NomListe, nil, nil, '');
    //on doit recharger les grilles !!!
    FreeAndNil(GGrilleLigFac);
    InitGrilleLigFac;
    GGrilleLigFac.ChargementGrille;
  end;
  
end;

Initialization
  registerclasses ( [ TOF_BTSUIVICDE ] ) ;
end.

