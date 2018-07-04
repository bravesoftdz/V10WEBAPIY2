{***********UNITE*************************************************
Auteur  ...... : DECOSSE
Créé le ...... : 29/08/2001
Modifié le ... : 27/07/2005
Description .. : Source des fonctions Zoom Ecriture depuis Pièce et Pièce
Suite ........ : depuis Ecriture
Suite ........ : 
Suite ........ : JP 27/07/05 / FQ 16172 : Suppression de la directive 
Suite ........ : IFNDEF COMPTA qui ne permettait plus l'accès à la
Suite ........ : consultation des pièces GC
Suite ........ :
Mots clefs ... : PIECEVERSECR;ECRVERSPIECE;
*****************************************************************}
Unit ECRPIECE_TOF ;

interface

Uses
    StdCtrls,
    Controls,
    Classes,
    forms,
    sysutils,
    ComCtrls,
    SaisComm,
    HEnt1,
{$IFDEF EAGLCLIENT}
    MainEAGL,
{$ELSE}
    Fe_Main,
   {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
{$ENDIF}
     menus,
    HCtrls,
    UTOF,
    Saisutil,
    vierge,
    UTOB,
    Windows,
  variants,
    EntGC,
    UtilPGI,
    ParamSoc
    ,uEntCommun
    ;

Procedure LanceZoomBordereau ( YY : RMVT; Totale : boolean=false) ;
Procedure LanceZoomPieceGC ( RefGC : String ) ;
procedure LanceZoomFrais(RefFrais, RefAff: String);

Type
  TOF_PIECEVERSECR = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
  private
    LesColonnes : string ;
    GS : THGRID ;
    XX     : RMVT ;
    Totale : boolean;
  public
  end ;

Type
  TOF_ECRVERSPIECE = Class (TOF)
    procedure OnClose; override;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
  private
        LesColonnes : string ;
        GS : THGRID ;
        XX : R_CLEDOC ;
        TobPiece : Tob;
        TobPort : Tob;
        bPiecePrec : TMenuItem;
        bPieceSuiv : TMenuItem;
        bPieceOrig : TMenuItem;
    procedure bPiecePrecClick (Sender: TObject);
    procedure bPieceSuivClick (Sender: TObject);
    procedure bPieceOrigClick (Sender: TObject);
  public
  end ;

  TOF_ECRVERSFRAIS = class(TOF)
    procedure OnClose; override;
    procedure OnLoad; override;
    procedure OnArgument(S: String ); override;
  private
    TOBActivite: Tob;
    Reference,
    LesColonnes: String;
    GS: THGRID;
    Largeur: Integer;
  end;
implementation

Procedure LanceZoomBordereau ( YY : RMVT; Totale : boolean=false) ;
Var S : String ;
begin
  if Totale then
    S := 'TOTAL="X";'
    else
    S := '';
  S := S + EncodeLC(YY) ;
  AGLLanceFiche('YY','YYPIECEVERSECR','','',s);
END ;

procedure TOF_PIECEVERSECR.OnLoad ;
var Q : TQuery ;
    SQL : String ;
    TOBEcr,TOBLig : TOB ;
    WhereEcr, LesChampsEcr, RefGescom : string;
    TobDesEcr : TOB;
begin
  inherited ;
  TOBEcr:=TOB.Create('',Nil,-1) ;
  TobDesEcr := TOB.Create('', nil, -1);
  if XX.Simul='' then
    XX.Simul:='N' ;
  WhereEcr := ' E_EXERCICE="'+XX.Exo+'" AND E_JOURNAL="' + XX.Jal + '" AND E_QUALIFPIECE="' + XX.Simul + '" AND E_NUMEROPIECE=' +IntToStr(XX.Num);
  if Totale then
    LesChampsEcr := 'E_JOURNAL, E_NUMEROPIECE, E_GENERAL, E_AUXILIAIRE, E_REFINTERNE, E_LIBELLE, E_DEBIT, E_CREDIT'
                  + ', E_ETABLISSEMENT, E_NUMLIGNE'
    else
    LesChampsEcr := 'E_NUMLIGNE, E_DATECOMPTABLE, E_GENERAL, E_AUXILIAIRE, E_REFINTERNE, E_LIBELLE, E_DEBIT, E_CREDIT, E_JOURNAL '
                  + ', E_ETABLISSEMENT, E_NUMEROPIECE';
  { Si on veut tout voir, recherche d'abord le E_REFGESCOM pour trouver ensuite
    toutes les écritures qui pointent sur le même
    ATTENTION !! peut être long (champs non indexé) }
  if Totale then
  begin
    Sql := 'SELECT E_REFGESCOM FROM ECRITURE WHERE ' + WhereEcr;
    Q := OpenSQL(Sql, True,-1,'',true);
    try
      if not Q.Eof then
        RefGescom := Q.FindField('E_REFGESCOM').AsString
        else
        RefGescom := '';
    finally
      Ferme(Q);
    end;
    if RefGescom <> '' then
    begin
      Sql := 'SELECT ' + LesChampsEcr
           + ' FROM ECRITURE WHERE E_REFGESCOM = "' + RefGescom + '"'
           + ' ORDER BY E_NUMEROPIECE DESC, E_NUMLIGNE' ;
      Q := OpenSQL(Sql, True,-1,'',true);
      try
        if Not Q.EOF then
        begin
          TOBEcr.LoadDetailDB('ECRITURE','','',Q,False,True) ;
          TOBEcr.Detail.Sort('E_JOURNAL;E_NUMEROPIECE;E_NUMLIGNE');
          TOBEcr.PutGridDetail(GS,False,False,LesColonnes,True) ;
        end;
      finally
        Ferme(Q);
      end;
    end;
  end else
  begin
    SQL:='SELECT ' + LesChampsEcr
       + ' FROM ECRITURE WHERE ' + WhereEcr
       + ' ORDER BY E_NUMLIGNE' ;
    Q:=OpenSQL(SQL,True,-1,'',true);
    if Not Q.EOF then
    BEGIN
      TOBEcr.LoadDetailDB('ECRITURE','','',Q,False,True) ;
      TOBEcr.PutGridDetail(GS,False,False,LesColonnes,True) ;
    END ;
    Ferme(Q);
  end;
  if TOBEcr.Detail.Count>0 then
  BEGIN
    TOBLig:=TOBEcr.Detail[0] ;
    if Totale then
    begin
      SetControlProperty('E_NUMEROPIECE', 'Visible', false);
      SetControlProperty('TE_NUMEROPIECE', 'Visible', false);
      SetControlProperty('E_JOURNAL', 'Visible', false);
      SetControlProperty('TE_JOURNAL', 'Visible', false);
    end else
    begin
      SetControlText('E_NUMEROPIECE', IntToStr(TOBLig.GetInteger('E_NUMEROPIECE')));
      SetControlText('E_JOURNAL', TOBLig.GetString('E_JOURNAL'));
    end;
    SetControlText('E_ETABLISSEMENT' ,TOBLig.GetString('E_ETABLISSEMENT'));
  END ;
  if assigned(TOBEcr) then FreeAndNil(TOBEcr);
  if assigned(TobDesEcr) then FreeAndNil(TobDesEcr);
end ;

procedure TOF_PIECEVERSECR.OnArgument (S : String ) ;
var ColTitre: HTStringList;
    NumDebit, NumCredit : integer;
begin
  Inherited ;
  ColTitre := HTStringList.Create;
  Totale := (pos('TOTAL="X"', S) > 0);
  if Totale then
  begin
    ReadtokenST(S);
    ColTitre.Add ('Jal;C;');
    ColTitre.Add ('Pce;C;');
  end else
  begin
    ColTitre.Add ('Lig;C;');
  end;
  ColTitre.Add ('Date;G;');
  ColTitre.Add ('Général;G;');
  ColTitre.Add ('Auxiliaire;G;');
  ColTitre.Add ('Référence;G;');
  ColTitre.Add ('Libellé;G;');
  ColTitre.Add ('Débit;D;');
  ColTitre.Add ('Crédit;D;');
  XX:=DecodeLC(S) ;
  GS:=THGRID(GetControl('FLIGECT')) ;
  if Totale then
  begin
    LesColonnes:='E_JOURNAL;E_NUMEROPIECE;E_DATECOMPTABLE;E_GENERAL;E_AUXILIAIRE;E_REFINTERNE;E_LIBELLE;E_DEBIT;E_CREDIT';
    GS.ColCount := 9;
    { Resize des colonnes }
    GS.ColWidths[0] := GS.ColWidths[0] + 15;
    GS.ColWidths[1] := 50;
    GS.ColWidths[4] := GS.ColWidths[4] - 20;
    GS.ColWidths[5] := GS.ColWidths[4] - 50;
    GS.ColFormats[2]:='###0' ;
    NumDebit := 7;
    NumCredit := 8;
  end else
  begin
    LesColonnes:='E_NUMLIGNE;E_DATECOMPTABLE;E_GENERAL;E_AUXILIAIRE;E_REFINTERNE;E_LIBELLE;E_DEBIT;E_CREDIT';
    GS.ColCount := 8;
    NumDebit := 6;
    NumCredit := 7;
  end;
  GS.Titres := ColTitre;
  GS.ColFormats[NumDebit]:='#,##0.00' ;
  GS.ColFormats[NumCredit]:='#,##0.00' ;
  ColTitre.Clear;
  FreeAndNil(ColTitre);
  TFVierge(Ecran).HMTrad.ResizeGridColumns(GS);
end ;

////////////////////////////////////////////////////////////////////
Procedure LanceZoomPieceGC ( RefGC : String ) ;
BEGIN
  AGLLanceFiche('YY','YYECRVERSPIECE','','',RefGC) ;
END ;

procedure LanceZoomFrais(RefFrais, RefAff: String);
var
  S: String;
begin
  // BDU - 16/11/06, si le champs E_REFGESCOM commence par l'un des préfixes (AA, AD ou AR)
  // on appel la visualisation des frais en ajoutant à la valeur du champ E_REFGESCOM
  // la valeur du champ E_AFFAIRE précédé d'un # de séparation
  S := Copy(RefFrais, 1, 3);
  if (S = 'AA;') or (S = 'AD;') or (S = 'AR;') then
    AGLLanceFiche('YY', 'YYECRVERSFRAIS', '', '', RefFrais + '#' + RefAff)
end;

procedure TOF_ECRVERSPIECE.OnClose ;
begin
  if assigned(TobPiece) then FreeAndNil(TobPiece);
  if assigned(TobPort) then FreeAndNil(TobPort);
end;

procedure TOF_ECRVERSPIECE.OnLoad ;
var TSql : TQuery;
    SQL : String;
begin
  inherited;
  TOBPiece := TOB.Create('',Nil,-1);
  TOBPort := TOB.Create('',Nil,-1);

  // BPY le 09/12/2003 => Fiche 12981
  {Pièce}
  SQL := 'SELECT * FROM PIECE WHERE GP_NATUREPIECEG="'+XX.NaturePiece+'" AND GP_SOUCHE="'+XX.Souche+'" AND GP_NUMERO='+IntToStr(XX.NumeroPiece)+' AND GP_INDICEG='+IntToStr(XX.Indice);
  TSql := OpenSQL (SQL,True,-1,'',true);
  TOBPiece.SelectDB ('', TSql);
  Ferme (TSql);

  Sql := 'SELECT Sum(GPT_TOTALHT) AS SUMPORT FROM PIEDPORT WHERE GPT_NATUREPIECEG="'+XX.NaturePiece+'" AND ' +
            'GPT_SOUCHE="'+XX.Souche+'" AND GPT_NUMERO='+IntToStr(XX.NumeroPiece)+' AND GPT_INDICEG='+IntToStr(XX.Indice);
  TSql := OpenSQL (SQL,True,-1,'',true);
  TOBPort.SelectDB ('', TSql);
  Ferme (TSql);

  if (TOBPiece.GetString ('GP_NATUREPIECEG') <> '') then
    SetControlText('FTITREPIECE',RechDom('GCNATUREPIECEG',TOBPiece.GetString ('GP_NATUREPIECEG') ,False));
  if (TOBPiece.GetString ('GP_TIERS') <> '') then
  begin
    SetControlText('FNOMTIERS',RechDom('GCTIERS',TOBPiece.GetString ('GP_TIERS'),False));
    SetControlText('GP_TIERS',TOBPiece.GetString ('GP_TIERS'));
  end;
  SetControlText('GP_NUMERO', TOBPiece.GetString ('GP_NUMERO'));
  SetControlText('GP_DATEPIECE', TOBPiece.GetString ('GP_DATEPIECE'));
  SetControlText('GP_REFINTERNE',TOBPiece.GetString ('GP_REFINTERNE'));
  THNumedit(Ecran.FindComponent('GP_ESCOMPTE')).Value := TOBPiece.GetInteger ('GP_ESCOMPTE');
  THNumedit(Ecran.FindComponent('GP_REMISEPIED')).Value := TOBPiece.GetInteger ('GP_REMISEPIED');

  THNumedit(Ecran.FindComponent('GP_TOTALHT')).Value := TOBPiece.GetDouble ('GP_TOTALHT');

  THNumedit(Ecran.FindComponent('GP_TOTALTTC')).Value := TOBPiece.GetDouble ('GP_TOTALTTC');

  THNumedit(Ecran.FindComponent('FTAXES')).Value := TOBPiece.GetDouble ('GP_TOTALTTC') - TOBPiece.GetDouble ('GP_TOTALHT');

  THNumEdit(Ecran.FindComponent('FPORT')).Value := TobPort.GetDouble ('SUMPORT');

  {Lignes}
  SQL := 'SELECT * FROM LIGNE WHERE GL_NATUREPIECEG="'+XX.NaturePiece+'" AND GL_SOUCHE="'+XX.Souche+'" AND GL_NUMERO='+IntToStr(XX.NumeroPiece)+' AND GL_INDICEG='+IntToStr(XX.Indice)+' ORDER BY GL_NUMLIGNE';
  TOBPiece.LoadDetailDBFromSQL('LIGNE',SQL);
  TOBPiece.PutGridDetail(GS,False,False,LesColonnes,True) ;
  GS.RowCount := GS.RowCount + 20;
  // fin BPY
end;

procedure TOF_ECRVERSPIECE.OnArgument (S : String ) ;
begin
  Inherited ;
  {$IFNDEF COMSX}
  XX:=DecodeRefCPGescom(S) ;
  {$ENDIF}
  LesColonnes:='GL_NUMLIGNE;GL_REFARTSAISIE;GL_LIBELLE;GL_QTEFACT;GL_PUHT;GL_REMISELIGNE;GL_MONTANTHT' ;
  GS:=THGRID(GetControl('FLIGGC')) ;
  GS.ColFormats[3]:='#,##0.00' ; GS.ColFormats[4]:='#,##0.00' ;
  GS.ColFormats[5]:='#,##0.00' ; GS.ColFormats[6]:='#,##0.00' ;
  bPiecePrec := TMenuItem(GetControl('mnPiecePrec'));
  bPieceSuiv := TMenuItem(GetControl('mnPieceSuiv'));
  bPieceOrig := TMenuItem(GetControl('mnPieceOrig'));
  if bPiecePrec<>nil then bPiecePrec.OnClick := bPiecePrecClick;
  if bPieceSuiv<>nil then bPieceSuiv.OnClick := bPieceSuivClick;
  if bPieceOrig<>nil then bPieceOrig.OnClick := bPieceOrigClick;
  end ;

procedure TOF_ECRVERSPIECE.bPieceSuivClick (Sender: TObject);
{$IFNDEF COMSX}
var CleDoc : R_CleDoc;
{$ENDIF}
begin
{$IFNDEF COMSX}
  if TOBPiece.GetValue ('GP_DEVENIRPIECE') <> '' then
  begin
    DecodeRefPiece (TobPiece.GetValue ('GP_DEVENIRPIECE'), CleDoc);
    AglLanceFiche ('YY','YYECRVERSPIECE','','',
                   CleDoc.NaturePiece + ';' + CleDoc.Souche + ';' +
                    FormatDateTime('ddmmyyyy', CleDoc.DatePiece) + ';' + IntToStr (CleDoc.NumeroPiece) +
                    ';' + IntToStr (CleDoc.Indice) + ';');
  end;
{$ENDIF}
end;

procedure TOF_ECRVERSPIECE.bPiecePrecClick (Sender: TObject);
{$IFNDEF COMSX}
var CleDoc : R_CleDoc;
{$ENDIF}
begin
{$IFNDEF COMSX}
  if ((GS.Row <= 0) or (GS.Row > TOBPiece.Detail.Count)) then Exit;
  if TOBPiece.Detail[GS.Row - 1].GetValue ('GL_PIECEPRECEDENTE') <> '' then
  begin
    DecodeRefPiece (TOBPiece.Detail[GS.Row - 1].GetValue ('GL_PIECEPRECEDENTE'), CleDoc);
    AglLanceFiche ('YY','YYECRVERSPIECE','','',
                    CleDoc.NaturePiece + ';' + CleDoc.Souche + ';' +
                    FormatDateTime('ddmmyyyy', CleDoc.DatePiece) + ';' + IntToStr (CleDoc.NumeroPiece) +
                    ';' + IntToStr (CleDoc.Indice) + ';');
  end;
{$ENDIF}
end;

procedure TOF_ECRVERSPIECE.bPieceOrigClick (Sender: TObject);
{$IFNDEF COMSX}
var CleDoc : R_CleDoc;
{$ENDIF}
begin
{$IFNDEF COMSX}
  if ((GS.Row <= 0) or (GS.Row > TOBPiece.Detail.Count)) then Exit;
  if TOBPiece.Detail[GS.Row - 1].GetValue ('GL_PIECEORIGINE') <> '' then
  begin
    DecodeRefPiece (TOBPiece.Detail[GS.Row - 1].GetValue ('GL_PIECEORIGINE'), CleDoc);
    AglLanceFiche ('YY','YYECRVERSPIECE','','',
                    CleDoc.NaturePiece + ';' + CleDoc.Souche + ';' +
                    FormatDateTime('ddmmyyyy', CleDoc.DatePiece) + ';' + IntToStr (CleDoc.NumeroPiece) +
                    ';' + IntToStr (CleDoc.Indice) + ';');
  end;
{$ENDIF}
end;

procedure TOF_ECRVERSFRAIS.OnArgument(S: String);
begin
  inherited;
  Reference := S;
  if GetParamSocSecur('SO_AFPRIXGEREFRAISACT', '') = 'PR' then
    LesColonnes := 'ACT_DATEACTIVITE;ACT_AFFAIRE;ACT_RESSOURCE;ACT_CODEARTICLE;' +
      'ACT_LIBELLE;ACT_QTEFAC;ACT_TOTPRCHARGE;ACT_MONTANTTVA;ACT_MONTANTTTC'
  else
    LesColonnes := 'ACT_DATEACTIVITE;ACT_AFFAIRE;ACT_RESSOURCE;ACT_CODEARTICLE;' +
      'ACT_LIBELLE;ACT_QTEFAC;ACT_TOTVENTE;ACT_MONTANTTVA;ACT_MONTANTTTC';
  GS := THGRID(GetControl('GRILLE')) ;
  GS.ColFormats[5] := '#,##0.00';
  GS.ColFormats[6] := '#,##0.00';
  GS.ColFormats[7] := '#,##0.00';
  GS.ColFormats[8] := '#,##0.00';
  Largeur := 570;
  GS.ColWidths[0] := 70;
  // Si regroupement des frais par affaire, pas d'affichage de la colonne AFFAIRE
  if Copy(Reference, 1, 2) = 'AA' then
    GS.ColWidths[1] := 0
  else
  begin
    GS.ColWidths[1] := 150;
    Inc(Largeur, 150);
  end;
  // La colonne RESSOURCE est affichée seulement en génération détaillée
  if Copy(Reference, 1, 2) = 'AD' then
  begin
    GS.ColWidths[2] := 70;
    Inc(Largeur, 70);
  end
  else
    GS.ColWidths[2] := 0;
  GS.ColWidths[3] := 70;
  GS.ColWidths[4] := 150;
  GS.ColWidths[5] := 70;
  GS.ColWidths[6] := 70;
  GS.ColWidths[7] := 70;
  GS.ColWidths[8] := 70;
end;

procedure TOF_ECRVERSFRAIS.OnClose;
begin
  inherited;
  if Assigned(TobActivite) then
    FreeAndNil(TobActivite);
end;

procedure TOF_ECRVERSFRAIS.OnLoad;
var
  Prefixe,
  SQL1,
  SQL2,
  Ref,
  Aff: String;
  Q: TQuery;
  TobSomme: TOB;
begin
  inherited;
  TobActivite := Tob.Create('les lignes frais', nil, -1);
  // Le préfixe est composé des deux premiers caractères
  Prefixe := Copy(Reference, 1, 2);
  // Le code affaire suit le caractère #
  Aff := Copy(Reference, Pos('#', Reference) + 1, Length(Reference));
  Ref := Copy(Reference, 4, Pos('#', Reference) - 4);
  // Sélection sur la référence comptable
  SQL1 := 'SELECT * FROM ACTIVITE WHERE ACT_REFCOMPTABLE = "' + Ref + '" ';
  // Si génération détaillée ou regroupement par affaire, sélection complémentaire sur l'affaire
  // Pour le regroupement par AFFAIRE/RESSOURCE, il manque un index. A voir selon besoin client
  if (Prefixe = 'AD') or (Prefixe = 'AA') then
    SQL2 := 'AND ACT_AFFAIRE = "' + Aff + '"';
  SQL1 := SQL1 + SQL2;
  TOBActivite.LoadDetailDBFromSQL('LIGNE', SQL1);
  TOBActivite.PutGridDetail(GS, False, False, LesColonnes, True);
  GS.RowCount := GS.RowCount;
  // Charge les cumuls
  if GetParamSocSecur('SO_AFPRIXGEREFRAISACT', '') = 'PR' then
    SQL1 := 'SELECT SUM(ACT_TOTPRCHARGE) SHT, SUM(ACT_MONTANTTTC) STTC, SUM(ACT_MONTANTTVA) STVA ' +
      'FROM ACTIVITE WHERE ACT_REFCOMPTABLE = "' + Ref + '" '
  else
    SQL1 := 'SELECT SUM(ACT_TOTVENTE) SHT, SUM(ACT_MONTANTTTC) STTC, SUM(ACT_MONTANTTVA) STVA ' +
      'FROM ACTIVITE WHERE ACT_REFCOMPTABLE = "' + Ref + '" ';
  SQL1 := SQL1 + SQL2;
  Q := OpenSQL(SQL1, True,-1,'',true);
  try
    TobSomme := Tob.Create('', nil, -1);
    try
      TobSomme.SelectDB('les sommes', Q);
      THNumedit(GetControl('SHT')).Value := TobSomme.GetDouble('SHT');
      THNumedit(GetControl('STVA')).Value := TobSomme.GetDouble('STVA');
      THNumedit(GetControl('STTC')).Value := TobSomme.GetDouble('STTC');
    finally
      FreeAndNil(TobSomme);
    end;
  finally
    Ferme(Q);
  end;
  Ecran.Width := Largeur;
end;
Initialization
  registerclasses([TOF_PIECEVERSECR,TOF_ECRVERSPIECE, TOF_ECRVERSFRAIS]) ;
end.
