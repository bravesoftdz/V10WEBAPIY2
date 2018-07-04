unit UtilConso;

interface
uses {$IFDEF VER150} variants,{$ENDIF} HEnt1,
     UTOB,
     Ent1,
     LookUp,
     SysUtils,
     UtilPGI,
     AGLInit,
     HCtrls,
{$IFDEF BTP}
     CalcOLEGenericBTP,
{$ENDIF}
{$IFNDEF EAGLCLIENT}
     FE_Main,
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     Maineagl,
     eMul,
     uTob,
{$ENDIF}
     EntGC,
     Classes,
     HMsgBox,
     ParamSoc,
     StockUtil,
     FactTob,
     UPlannifchUtil,uEntCommun;

Type TTraitConso = (TTcoCreat, TTcoModif, TTcoDelete, TTcoTransform,TTcoReGroup,TtcoLivraison);

type TGestionConso = class
  private
    TOBConso : TOB;
    TOBOldConso : TOB;
    TOBReceptionFou : TOB;
    TOBReceptHorsLien : TOB;
    GestionConso : TTraitConso;
    procedure MajoldConso;
    procedure DefiniQteLivree(TOBL: TOB; QteLivre: double; LienVente : double);
    procedure ReajustePrixAchat;
    function 	IsLivraisonClient (TOBL : TOB) : boolean;
    procedure FlagRetourChantierAssocie (NumConsoRecep : double;LienVente : double);
		procedure PositionneEventuelDoublon (TOBCourante : TOB;Qte : double);
    procedure AffecteIndices;
    procedure RecupCoefsPO(TOBLigne: TOB; venteAchat : string; var CoefpaPr, CoefPrPv: double);
    procedure CreeAssociation(TOBPhases, TOBL: TOB;  Nummouv : double = 0);
    function LibellePiece(TOBDR: TOB): string;

  public
    constructor create;
    destructor destroy ; override;
    procedure Clear;
    Function CreerConso(TobLigne :TOB; Numphase : string; Mode: TTraitConso; TOBDest:TOB=nil) : Double;
    function GetTOBrecepHLien: TOB;
    Procedure MajTableConso;
    procedure MemoriseLienReception(TOBL: TOB; TOBReception : TOB);
    procedure ReajusteReception;
    procedure recupereReceptions (TOBL : TOB);
    procedure recupereReceptionsHorsLien(TOBPiece,TOBAffaire: TOB); overload;
		procedure recupereReceptionsHorsLien (TOBPiece : TOB;Laffaire : String; Larticle : string=''); overload;
    procedure InitReceptions;
    function NbrReceptionsHorsLien : integer;
    procedure DefiniReceptionsFromHlienAssocie(TOBL, TOBDetRec: TOB);
    function GetQteLivrable(TOBL: TOB; WithGeneration : boolean= false): double;
    procedure RepertorieReceptionsHLiensFromLIV(TOBPiece: TOB;Affaire: String);
    procedure GestionModifLivraison(TOBGenere, TOBCONSODEL: TOB);
end;

Procedure AnnuleConso(TobPiece : TOB; GestionConso : TTraitConso = TTcoDelete);
function PieceAutoriseConso (Naturepiece : string) : boolean;
procedure ReajusteRecptFromLivr(TOBL: TOB);
function NewConsoFromLigne (TOBLigne,TOBC : TOB) : boolean;
function EncodeRefPieceLivr (TOBl : TOB) : string; overload;
function EncodeRefPieceLivr (refpieceStd : string) : string; overload;
implementation

uses BTPUtil,UtilSaisieConso,FactUtil,FactComm,UspecifPOC;

function PieceAutoriseConso (Naturepiece : string) : boolean;
begin
  result := false;
  if (Naturepiece = 'CF') OR (Naturepiece = 'BLF') OR
     (Naturepiece = 'BLF') or (Naturepiece = 'LBT') OR
     (Naturepiece = 'CFR') or (Naturepiece = 'LFR') OR
     (Naturepiece = 'BFC') OR (Naturepiece = 'BFA') OR
     (NaturePiece = 'AF') OR (NaturePiece = 'AFS') OR
     (Naturepiece = 'FF') OR (Naturepiece = 'PBT') OR
     (Naturepiece = 'BCM') OR (Naturepiece = 'BLC') OR
     ((Naturepiece='BCE') AND (VH_GC.BTCODESPECIF = '001'))
     then result := true;
end;

procedure ReAjusteConsoPrec (TOBPiece : TOB);
var SQL : String;
    QQ : TQuery;
    TOBConso,TOBC,TOBOldConso,TOBOC,TOBL : TOB;
    // Conso des réceptionsd fourn
    TOBCOnsoFou : TOB;
    Indice : integer;
    Naturepiece : string;
begin
  TOBConso := TOB.Create ('LES CONSO', nil,-1);
  TOBOldConso := TOB.Create ('LES ANCIENNES CONSO',nil,-1);
  TOBConsoFou := TOB.Create ('LES RECEPTIONS FOU',nil,-1);
  Naturepiece := TOBPiece.getvalue('GP_NATUREPIECEG');

  Sql := 'SELECT * FROM CONSOMMATIONS WHERE BCO_NATUREPIECEG="'+
          Naturepiece+'" AND BCO_NUMERO='+
          IntToStr(TOBPiece.GetValue('GP_NUMERO'))+ ' AND BCO_INDICEG='+
          IntToStr(TOBPiece.GetValue('GP_INDICEG')) + ' AND BCO_LIENTRANSFORME<>0';
  QQ := OpenSql (sql,true);
  TRY
    if not QQ.eof then
    begin
      TOBConso.loaddetaildb ('CONSOMMATIONS','','',QQ,false);
      for Indice := 0 to TOBConso.detail.count -1 do
      begin
        ferme (QQ);
        TOBC := TOBConso.detail[indice];
        QQ := OpenSql ('SELECT * FROM CONSOMMATIONS WHERE BCO_NUMMOUV='+FloatToStr(TOBC.GetValue('BCO_LIENTRANSFORME'))+' AND BCO_INDICE=0',true);
        if not QQ.eof then
        begin
          TOBOC := TOB.Create ('CONSOMMATIONS',TobOldConso,-1);
          TOBOC.SelectDB ('',QQ);
          TOBOC.PutValue('BCO_TRANSFORME','-');
          TOBOC.PutValue('BCO_QUANTITE',TOBOC.GetValue('BCO_QUANTITE')+TOBC.GetValue('BCO_QUANTITE'));
          calculeLaLigne (TOBOC);
        end;
      end;
      if TOBOldConso.detail.count > 0 then
      begin
        if not TOBOldConso.UpdateDB(false) then V_PGI.IOError:=OeUnknown;
      end;
    end;

    if (NaturePiece = 'LBT') or (NaturePiece = 'BLC') then
    begin
      for Indice := 0 to TOBPiece.detail.count -1 do
      begin
        TOBL := TOBPiece.detail[Indice];
        if TOBL.GetValue('BLP_NUMMOUV') > 0 then
        begin
          ReajusteRecptFromLivr (TOBL);
        end;
      end;
    end;

  FINALLY
    ferme (QQ);
    TOBConso.free;
    TOBOldConso.free;
    TOBConsoFou.free;
  END;
end;

{***********A.G.L.***********************************************
Auteur  ...... : franck Vautrain
Créé le ...... : 02/03/2004
Modifié le ... :   /  /
Description .. : Annulation des conso
Mots clefs ... :
*****************************************************************}
Procedure AnnuleConso(TobPiece : TOB; GestionConso : TTraitConso = TTcoDelete);
var SQL : String;
Begin
  if (GestionConso <> TTcoTransform) and
     (gestionConso <> TtcoRegroup) then ReAjusteConsoPrec (TOBPiece);
  //
  if V_PGI.IOError = OeOk then
  begin
    if TOBpiece.getString('GP_NATUREPIECEG')='FAC' then
    begin
      Sql := 'UPDATE CONSOMMATIONS '+
             'SET BCO_NATUREPIECEG="",BCO_SOUCHE="",BCO_NUMERO=0,BCO_INDICEG=0  '+
             'WHERE '+
             'BCO_NATUREPIECEG="'+TOBPiece.getvalue('GP_NATUREPIECEG')+'" AND '+
             'BCO_NUMERO='+IntToStr(TOBPiece.GetValue('GP_NUMERO'))+' AND '+
             'BCO_INDICEG='+ IntToStr(TOBPiece.GetValue('GP_INDICEG'));
    end else
    begin
      Sql := 'DELETE FROM CONSOMMATIONS WHERE BCO_NATUREPIECEG="'+
              TOBPiece.getvalue('GP_NATUREPIECEG')+'" AND BCO_NUMERO='+
              IntToStr(TOBPiece.GetValue('GP_NUMERO'))+ ' AND BCO_INDICEG='+
              IntToStr(TOBPiece.GetValue('GP_INDICEG'));
    end;
    ExecuteSQL (SQL);
  end;
End;

{ TGestionConso }

constructor TGestionConso.create;
begin
 TOBConso := TOB.create('Phase Conso',nil,-1);
 TOBOldConso := TOB.create('les Conso precedentes',nil,-1);
 TOBReceptionFou := TOB.create('les Receptions fourniseurs',nil,-1);
 TOBReceptHorsLien := TOB.Create ('LES recep hors Liens',nil,-1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : franck Vautrain
Créé le ...... : 02/03/2004
Modifié le ... : 02/03/2004
Description .. : Création des consommations à partir du TOB Ligne
Mots clefs ... :
*****************************************************************}
Function TGestionConso.CreerConso(TobLigne :TOB; Numphase : string; Mode: TTraitConso;TOBDest:TOB=nil) : Double;
Var Part      : String;
    Part0     : String;
    Part1     : String;
    Part2     : String;
    Part3     : String;
    Part4     : String;
    Requete   : String;
    Annee     : word;
    Mois      : word;
    Jour      : word;
    Semaine   : integer;
    DateMouv  : TDateTime;
    Q         : TQuery;
    Nature    : String;
    NumMouv   : Double;
    TheRetour : TGncERROR;
    TOBC      : TOB;
    RatioVa,RatioVente : double;
    stprec,VenteAchat,TypeArticle,Select : string;
    Indice,II : integer;
    TOBC1,TOBDC1,TOBDETCONSO : TOB;
    CoefPAPR,COEfPrPV : double;
    cledoc : R_CLEDOC;
Begin
  result := -1;

  if TobLigne = nil then Exit;
  
  TOBDETCONSO := TOB.Create('LES CONSO',nil,-1);
  TRY
    if TobLigne.GetValue('GL_AFFAIRE') = '' then exit;
    GestionConso := Mode;
    VenteAchat := GetInfoParPiece(TobLigne.GetValue('GL_NATUREPIECEG'), 'GPP_VENTEACHAT');

    Part  := TobLigne.GetValue('GL_AFFAIRE');
    Part0 := '';
    Part1 := '';
    Part2 := '';
    Part3 := '';
    Part4 := TobLigne.GetValue('GL_AVENANT');

    DateMouv := TobLigne.GetValue('GL_DATEPIECE');

    // Découpage du code Affaire
  {$IFDEF BTP}
    BTPCodeAffaireDecoupe(Part,Part0,Part1,Part2,Part3, Part4, TaModif,false);
  {$ELSE}
    CodeAffaireDecoupe(Part,Part0,Part1,Part2,Part3, Part4, TaModif,false);
  {$ENDIF}

    // Découpage de la date
    DecodeDate(DateMouv, Annee, Mois, Jour);

    // Recherche du Numéro de semaine
    Semaine := NumSemaine(DateMouv);
    TypeArticle := TobLigne.GetValue('GL_TYPEARTICLE');
    // Recherche de la Nature du Mouvement si Prestation
    if TypeArticle = 'PRE' then
    Begin
      Requete := 'SELECT GA_ARTICLE,GA_LIBELLE,N1.BNP_TYPERESSOURCE,N1.BNP_LIBELLE FROM ARTICLE ART ' +
                 'LEFT JOIN NATUREPREST N1 ON ART.GA_NATUREPRES=N1.BNP_NATUREPRES ' +
                 'WHERE GA_TYPEARTICLE="PRE" AND GA_ARTICLE="'+TOBLigne.GetValue('GL_ARTICLE')+'"';
      Q := OpenSQL(Requete, True);
      if not Q.eof then
      begin
        Nature := Q.findfield('BNP_TYPERESSOURCE').asString ;
        ferme (Q);
      end else
      Begin
        ferme (Q);
        exit;
      end;
    end;

    if venteAchat = 'ACH' then
    begin
      CoefPaPr := 0;
      CoefPrPv := 0;
      RatioVente := GetRatio(TOBLigne, nil, trsVente);
      if TobLigne.getString('GL_NATUREPIECEG') = 'BCM' then
      begin
        CoefPAPR := 1;
        CoefPRPV := 1;
      end else
      begin
        RecupCoefsPO (TOBLigne,venteAchat,CoefPaPr,CoefPrPv);
      end;
      if VH_GC.BTCODESPECIF = '001' then GetCoefPoc (TobLigne.GetString('GL_AFFAIRE'),CoefPaPr,CoefPrPv);
    end else
    begin
      RatioVente := 1;
      CoefPaPr := 0;
      CoefPrPv := 0;
      if TOBLigne.GetValue('GL_DPA') <> 0 then CoefPaPR := TOBLigne.GetValue('GL_DPR')/TOBLigne.GetValue('GL_DPA');
      if TOBLigne.GetValue('GL_DPR') <> 0 then CoefPRPV := TOBLigne.GetValue('GL_PUHT')/TOBLigne.GetValue('GL_DPR');
    end;

    TOBC := nil;
    if (Mode = TTcoModif) and (Pos(TOBLigne.GetValue('GL_NATUREPIECEG'),GetPieceAchat (false,false,false,true))>0) then
    begin
      TOBDETCONSO.ClearDetail;
      if not TOBLigne.FieldExists ('BLP_NUMMOUV') then
      begin
        TOBLigne.AddChampSupValeur ('BLP_NUMMOUV',0,false);
      end else
      begin
        if VarIsNull(TOBLigne.getValue('BLP_NUMMOUV')) or (VarAsType(TOBLigne.getValue('BLP_NUMMOUV'), varString) = #0) then
          TOBLigne.PutValue('BLP_NUMMOUV',0);
      end;

      if TOBLigne.GetValue('BLP_NUMMOUV') <> 0 then
      begin
        // modification d'un reception fournisseur
        Select := 'SELECT * FROM CONSOMMATIONS WHERE BCO_NUMMOUV='+inttostr(TOBLigne.GetValue('BLP_NUMMOUV'))+' AND '+
                  'BCO_NATUREPIECEG="'+TOBLigne.GetValue('GL_NATUREPIECEG')+'" AND '+
                  'BCO_SOUCHE="'+TOBLigne.GetValue('GL_SOUCHE')+'" AND '+
                  'BCO_NUMERO='+InttoStr(TOBLigne.GetValue('GL_NUMERO'))+' AND '+
                  'BCO_INDICEG='+InttoStr(TOBLigne.GetValue('GL_INDICEG'))+' ORDER BY BCO_NUMMOUV,BCO_INDICE';
        Q := OpenSql (select,true);
        if not Q.eof then
        begin
          TOBDETCONSO.LoadDetailDB ('CONSOMMATIONS','','',Q,True);
          Repeat
            TOBC := TOBDETCONSO.Detail[II];
            NumMouv := TOBC.GetValue('BCO_NUMMOUV');
            TOBC.PutValue('BCO_AFFAIRE', TobLigne.GetValue('GL_AFFAIRE'));
            TOBC.PutValue('BCO_AFFAIRE0', Part0);
            TOBC.PutValue('BCO_AFFAIRE1', Part1);
            TOBC.PutValue('BCO_AFFAIRE2', Part2);
            TOBC.PutValue('BCO_AFFAIRE3', Part3);
            TOBC.PutValue('BCO_PHASETRA', NumPhase);
            TOBC.PutValue('BCO_MOIS', Mois);
            TOBC.PutValue('BCO_SEMAINE', Semaine);
            TOBC.PutValue('BCO_DATEMOUV', DateMouv);
            if TOBC.GetInteger('BCO_INDICE') = 0 then
            begin
              TOBC.PutValue('BCO_QUANTITE', TobLigne.GetValue('GL_QTERESTE') * RatioVente);
              TOBC.PutValue('BCO_DPA', TobLigne.GetValue('GL_PUHTNET')/Ratiovente/(TOBLigne.getValue('GL_PRIXPOURQTE')));
              TOBC.PutValue('BCO_DPA', Arrondi(TOBC.GEtValue('BCO_DPA'),V_PGI.okdecP));
              TOBC.PutValue('BCO_DPR', Arrondi(TOBC.GEtValue('BCO_DPA')*CoefPaPr,V_PGI.okdecP));
              TOBC.PutValue('BCO_PUHT',Arrondi(TOBC.GEtValue('BCO_DPR')*CoefPRPV,V_PGI.okdecP));
              CalculeLaLigne(TOBC);
              Result := NumMouv;
            end;
            TOBC.ChangeParent(TOBConso,-1);
          until II >= TOBDETCONSO.detail.count;
        end;
        ferme (Q);
      end;
    end;
    if (Mode = TTcoModif) and (Pos(TOBLigne.GetValue('GL_NATUREPIECEG'),'CF;CFR')>0) then
    begin
    	TOBDETCONSO.ClearDetail;
      if not TOBLigne.FieldExists ('BLP_NUMMOUV') then
      begin
        TOBLigne.AddChampSupValeur ('BLP_NUMMOUV',0,false);
      end else
      begin
        if VarIsNull(TOBLigne.getValue('BLP_NUMMOUV')) or (VarAsType(TOBLigne.getValue('BLP_NUMMOUV'), varString) = #0) then
          TOBLigne.PutValue('BLP_NUMMOUV',0);
      end;

      if TOBLigne.GetValue('BLP_NUMMOUV') <> 0 then
      begin
        // modification d'une commande fournisseur
        Select := 'SELECT * FROM CONSOMMATIONS WHERE BCO_NUMMOUV='+Floattostr(TOBLigne.GetValue('BLP_NUMMOUV'))+' AND '+
                  'BCO_NATUREPIECEG="'+TOBLigne.GetValue('GL_NATUREPIECEG')+'" AND '+
                  'BCO_SOUCHE="'+TOBLigne.GetValue('GL_SOUCHE')+'" AND '+
                  'BCO_NUMERO='+InttoStr(TOBLigne.GetValue('GL_NUMERO'))+' AND '+
                  'BCO_INDICEG='+InttoStr(TOBLigne.GetValue('GL_INDICEG'))+' ORDER BY BCO_NUMMOUV,BCO_INDICE';
        Q := OpenSql (select,true);
        if not Q.eof then
        begin
          TOBDETCONSO.LoadDetailDB ('CONSOMMATIONS','','',Q,True);
          repeat
            TOBC := TOBDETCONSO.detail[II];
            NumMouv := TOBC.GetValue('BCO_NUMMOUV');
            TOBC.PutValue('BCO_AFFAIRE', TobLigne.GetValue('GL_AFFAIRE'));
            TOBC.PutValue('BCO_AFFAIRE0', Part0);
            TOBC.PutValue('BCO_AFFAIRE1', Part1);
            TOBC.PutValue('BCO_AFFAIRE2', Part2);
            TOBC.PutValue('BCO_AFFAIRE3', Part3);
            TOBC.PutValue('BCO_PHASETRA', NumPhase);
            TOBC.PutValue('BCO_MOIS', Mois);
            TOBC.PutValue('BCO_SEMAINE', Semaine);
            TOBC.PutValue('BCO_DATEMOUV', DateMouv);
            if TOBC.GetInteger('BCO_INDICE')= 0 then
            begin
              TOBC.PutValue('BCO_QUANTITE', TobLigne.GetValue('GL_QTERESTE')* RatioVente);
              TOBC.PutValue('BCO_QTEINIT', TOBC.GetValue('BCO_QUANTITE'));
              TOBC.PutValue('BCO_DPA', TobLigne.GetValue('GL_PUHTNET')/RatioVente/(TOBLigne.getValue('GL_PRIXPOURQTE')));
              TOBC.PutValue('BCO_DPA', Arrondi(TOBC.GEtValue('BCO_DPA'),V_PGI.okdecP));
              TOBC.PutValue('BCO_DPR', Arrondi(TOBC.GEtValue('BCO_DPA')*CoefPaPr,V_PGI.okdecP));
              TOBC.PutValue('BCO_PUHT',Arrondi(TOBC.GEtValue('BCO_DPR')*CoefPRPV,V_PGI.okdecP));
              //
              if TOBC.getValue('BCO_QUANTITE')=0 then
              begin
                TOBC.Free;
                ferme (Q);
                result := 0;
                exit;
              end;
              CalculeLaLigne(TOBC);
              Result := NumMouv;
            end;
            TOBC.ChangeParent(TOBConso,-1)
          until II >= TOBDETCONSO.Detail.count;
        end;
        ferme (Q);
      end;
    end;

    if (TOBC = nil) then
    begin
      // Calcul du Numéro de Mouvement
      TheRetour := GetNumUniqueConso (NumMouv);
      if TheRetour = gncAbort then
      BEGIN
        V_PGI.IOError := oeUnknown;
        Exit;
      END;

      // Chargement de la TOB Conssomation
      if TOBDest = nil then
      begin
        TOBC := Tob.Create('CONSOMMATIONS', TobConso, -1);
      end else
      begin
        TOBC := Tob.Create('CONSOMMATIONS', TobDest, -1);
      end;

      TOBC.PutValue('BCO_AFFAIRE', TobLigne.GetValue('GL_AFFAIRE'));
      TOBC.PutValue('BCO_AFFAIRE0', Part0);
      TOBC.PutValue('BCO_AFFAIRE1', Part1);
      TOBC.PutValue('BCO_AFFAIRE2', Part2);
      TOBC.PutValue('BCO_AFFAIRE3', Part3);
      TOBC.PutValue('BCO_PHASETRA', NumPhase);
      TOBC.PutValue('BCO_MOIS', Mois);
      TOBC.PutValue('BCO_SEMAINE', Semaine);
      TOBC.PutValue('BCO_DATEMOUV', DateMouv);
    //INT,ST,AUT,LOC
      if TypeArticle = 'PRE' then
      begin
        if Nature = 'SAL' then
          TOBC.PutValue('BCO_NATUREMOUV', 'MO')
        else if  Nature = 'ST' then
          TOBC.PutValue('BCO_NATUREMOUV', 'EXT')
        else if  Nature = 'INT' then
          TOBC.PutValue('BCO_NATUREMOUV', 'EXT')
        else if  Nature = 'AUT' then
          TOBC.PutValue('BCO_NATUREMOUV', 'EXT')
        else if  Nature = 'LOC' then
          TOBC.PutValue('BCO_NATUREMOUV', 'EXT')
        else if  Nature = 'MAT' then
          TOBC.PutValue('BCO_NATUREMOUV', 'RES')
        else if  Nature = 'OUT' then
          TOBC.PutValue('BCO_NATUREMOUV', 'RES')
      end else if (TypeArticle = 'MAR') or (TypeArticle = 'ARP') then
      Begin
        TOBC.PutValue('BCO_NATUREMOUV', 'FOU');
      end else if TypeArticle = 'FRA' then
      Begin
        TOBC.PutValue('BCO_NATUREMOUV', 'FRS');
      end;

      TOBC.PutValue('BCO_LIBELLE', TobLigne.GetValue('GL_LIBELLE'));
      TOBC.PutValue('BCO_TRANSFORME', '-');
      StPrec := TOBLigne.GetValue('GL_PIECEPRECEDENTE');

      //Pas de gestion de la pièce précédente si celle-ci est une demande d'achat : MODIF BRL 20/01/2012
      if stprec <> '' then
      begin
        DecodeRefPiece (StPrec,Cledoc);
        if cledoc.NaturePiece = 'DEF' then StPrec := '';
      end;
      //

      if (stprec <> '') and ((gestionConso = TtcoTransform) or (GestionConso = TTcoRegroup)) then
      begin
        TOBC.PutValue('BCO_LIENTRANSFORME',TobLigne.GetValue('NUMCONSOPREC'));
        TOBC.PutValue('BCO_TRANSFORME',    TobLigne.GetValue('BCO_TRANSFORME'));

        //FS#2736 - DELABOUDINIERE - Consommation non créée suite à génération d'un avoir fournisseur issu d'un retour.
        IF (TOBC.GetVALUE('BCO_TRANSFORME') <> '-') OR (TOBC.GetVALUE('BCO_TRANSFORME') <> 'X') then TOBC.PutValue('BCO_TRANSFORME', '-');

        TOBC.PutValue('BCO_TRAITEVENTE',TobLigne.GetValue('BCO_TRAITEVENTE'));
        TOBC.PutValue('BCO_QTEVENTE',TobLigne.GetValue('BCO_QTEVENTE'));
        // inutile de récupérer le lienvente bicose a 0 sur les indices a 0
      end else
      if (stprec <> '') and ((gestionConso = TtcoModif) or (GestionConso = TtcoCreat)) then
      begin
        TOBC.PutValue('BCO_LIENTRANSFORME',TobLigne.GetValue('BCO_LIENTRANSFORME'));
        TOBC.PutValue('BCO_LIENRETOUR',    TobLigne.GetValue('BCO_LIENRETOUR'));
        TOBC.PutValue('BCO_TRANSFORME',    TobLigne.GetValue('BCO_TRANSFORME'));
        TOBC.PutValue('BCO_TRAITEVENTE',   TobLigne.GetValue('BCO_TRAITEVENTE'));
        TOBC.PutValue('BCO_LIENVENTE',     TobLigne.GetValue('BCO_LIENVENTE'));
        if (not IsRetourFournisseur (TOBLigne)) and (not IsRetourClient(TOBLigne)) then
        begin
          TOBC.PutValue('BCO_QTEVENTE',TobLigne.GetValue('BCO_QTEVENTE'));
          TOBC.PutValue('BCO_QTEINIT',TobLigne.GetValue('BCO_QTEINIT'));
        end;
      end else
      if ((gestionConso = TtcoModif) or (GestionConso = TtcoCreat)) and
          ((TOBC.GetValue('BCO_NATUREPIECEG')='CF') or (TOBC.GetValue('BCO_NATUREPIECEG')='CFR')) then
      begin
        TOBC.PutValue('BCO_QTEINIT',TobLigne.GetValue('GL_QTERELIQUAT') + TobLigne.GetValue('GL_QTERESTE'));
      end;

      TOBC.PutValue('BCO_NUMMOUV', NumMouv);
      TOBC.PutValue('BCO_INDICE',0);
      if TObLigne.getValue('GL_PRIXPOURQTE')= 0 then TOBLigne.putvalue('GL_PRIXPOURQTE',1);
      TOBC.PutValue('BCO_CODEARTICLE', TobLigne.GetValue('GL_CODEARTICLE'));
//      TOBC.PutValue('BCO_RESSOURCE', TobLigne.GetValue('GL_RESSOURCE'));
      TOBC.PutValue('BCO_ARTICLE', TobLigne.GetValue('GL_ARTICLE'));
      TOBC.PutValue('BCO_FAMILLENIV1', TobLigne.GetValue('GL_FAMILLENIV1'));
      TOBC.PutValue('BCO_FAMILLENIV2', TobLigne.GetValue('GL_FAMILLENIV2'));
      TOBC.PutValue('BCO_FAMILLENIV3', TobLigne.GetValue('GL_FAMILLENIV3'));
      if TobLigne.GetString('GL_NATUREPIECEG')='BCM' then
      begin
        TOBC.PutValue('BCO_QUANTITE', TobLigne.GetValue('GLC_QTEAPLANIF'));
        TOBC.PutValue('BCO_QUALIFQTEMOUV', 'H');
      end else
      begin
        TOBC.PutValue('BCO_QUANTITE', TobLigne.GetValue('GL_QTERESTE') * RatioVente);
        TOBC.PutValue('BCO_QUALIFQTEMOUV', TobLigne.GetValue('GL_QUALIFQTEVTE'));
      end;

      //FV1 : 23/08/2017 - FS#2648 - DELABOUDINIERE - Rendre les consos associées à un appel facturable
      if (TOBLigne.fieldExists('FACTURABLE')) then
        TOBC.PutValue('BCO_FACTURABLE', TOBLigne.GetValue('FACTURABLE'))
      else
        TOBC.PutValue('BCO_FACTURABLE', 'N');
      //

      if Mode = TtcoLivraison then
      begin
        if (TOBLigne.fieldExists ('DPARECUPFROMRECEP')) and (TobLigne.GetValue('DPARECUPFROMRECEP') > 0) then
          TobLigne.PutValue('GL_DPA',TobLigne.GetValue('DPARECUPFROMRECEP'));
        if Part0 = 'W' then
        begin
          // Appel d'intervention
          TOBC.PutValue('BCO_FACTURABLE', 'A');
          TOBC.putValue('BCO_AFFAIRESAISIE',GetContrat(TobLigne.GetValue('GL_AFFAIRE')));
          TOBC.putValue('BCO_QTEFACTUREE',TOBC.GetValue('BCO_QUANTITE'));
        end;
      end;

      if (venteAchat = 'ACH') then
      begin
        if TobLigne.GetString('GL_NATUREPIECEG')='BCM' then
        begin
          TOBC.PutValue('BCO_DPA', ARRONDI(TobLigne.GetValue('GL_TOTALHT')/TobLigne.GetValue('GLC_QTEAPLANIF'),V_PGI.okdecP));
        end else
        begin
          TOBC.PutValue('BCO_DPA', TobLigne.GetValue('GL_PUHTNET')/RatioVente/(TOBLigne.getValue('GL_PRIXPOURQTE')));
          TOBC.PutValue('BCO_DPA', Arrondi(TOBC.GEtValue('BCO_DPA'),V_PGI.okdecP));
        end;
        TOBC.PutValue('BCO_DPR', Arrondi(TOBC.GEtValue('BCO_DPA')*CoefPaPr,V_PGI.okdecP));
        TOBC.PutValue('BCO_PUHT',Arrondi(TOBC.GEtValue('BCO_DPR')*CoefPRPV,V_PGI.okdecP));
      end else
      begin
        TOBC.PutValue('BCO_DPA', (TobLigne.GetValue('GL_DPA')/(TOBLigne.getValue('GL_PRIXPOURQTE'))));
        TOBC.PutValue('BCO_DPR', (TobLigne.GetValue('GL_DPA')/(TOBLigne.getValue('GL_PRIXPOURQTE'))*CoefPaPr));
        TOBC.PutValue('BCO_PUHT', (TobLigne.GetValue('GL_DPR')/(TOBLigne.getValue('GL_PRIXPOURQTE'))*CoefPrPv));
        TOBC.PutValue('BCO_DPA', Arrondi(TOBC.GEtValue('BCO_DPA'),V_PGI.okdecP));
        TOBC.PutValue('BCO_DPR', Arrondi(TOBC.GEtValue('BCO_DPR'),V_PGI.okdecP));
        TOBC.PutValue('BCO_PUHT', Arrondi(TOBC.GEtValue('BCO_PUHT'),V_PGI.okdecP));
    (*
        TOBC.PutValue('BCO_DPR', TobLigne.GetValue('GL_DPR')/(TOBLigne.getValue('GL_PRIXPOURQTE')*RatioVente));
        TOBC.PutValue('BCO_PUHT', TobLigne.GetValue('GL_PUHT')/(TOBLigne.getValue('GL_PRIXPOURQTE')* RatioVente));
    *)
      end;

      TOBC.PutValue('BCO_VALIDE', '-');
      TOBC.PutValue('BCO_NATUREPIECEG', TobLigne.GetValue('GL_NATUREPIECEG'));
      TOBC.PutValue('BCO_SOUCHE', TobLigne.GetValue('GL_SOUCHE'));
      TOBC.PutValue('BCO_NUMERO', TobLigne.GetValue('GL_NUMERO'));
      TOBC.PutValue('BCO_INDICEG', TobLigne.GetValue('GL_INDICEG'));
      TOBC.PutValue('BCO_NUMORDRE', TobLigne.GetValue('GL_NUMORDRE'));
      if Mode = TtcoCreat then
      begin
        if (not IsRetourFournisseur (TOBLigne)) and (not IsRetourClient(TOBLigne)) then
        begin
          if TobLigne.GetString('GL_NATUREPIECEG')='BCM' then
          begin
            TOBC.PutValue('BCO_QTEINIT', TOBC.GetDouble('BCO_QUANTITE'));
          end else
          begin
            TOBC.PutValue('BCO_QTEINIT', TobLigne.GetValue('GL_QTERESTE') * RatioVente);
          end;
        end;
      end;
      if Mode = TtcoLivraison then DefiniQteLivree (TOBLigne,TOBC.GetValue('BCO_QUANTITE'),NumMouv);
    //OBC.PutValue('BCO_LIENTRANSFORME', TOBLigne.getValue('NUMCONSOPREC'));

      CalculeLaLigne(TOBC);

      Result := NumMouv;
      if (stprec <> '') and ((gestionConso = TtcoTransform) or (GestionConso = TTcoRegroup)) then
      begin
        // récup des éléments préalablement livrés
        Q := OpenSql ('SELECT * FROM CONSOMMATIONS WHERE BCO_NUMMOUV='+
                      FloatToStr(TobLigne.GetValue('NUMCONSOPREC'))+ ' AND BCO_INDICE > 0',True);
        if not Q.eof then
        begin
          TOBDC1 := TOB.Create ('LES CONSO DET PREC',nil,-1);
          TOBDC1.loadDetailDb ('CONSOMMATIONS','','',Q,false);
          for indice := 0 to TOBDC1.detail.count -1 do
          begin
            TOBC1 := Tob.Create('CONSOMMATIONS', TobConso, -1);
            TOBC1.Dupliquer (TOBDC1.detail[Indice],false,true);
            TOBC1.PutValue('BCO_NUMMOUV', NumMouv);
            TOBC1.PutValue('BCO_LIENTRANSFORME',TobLigne.GetValue('NUMCONSOPREC'));
            TOBC1.PutValue('BCO_NATUREPIECEG', TobLigne.GetValue('GL_NATUREPIECEG'));
            TOBC1.PutValue('BCO_SOUCHE', TobLigne.GetValue('GL_SOUCHE'));
            TOBC1.PutValue('BCO_NUMERO', TobLigne.GetValue('GL_NUMERO'));
            TOBC1.PutValue('BCO_INDICEG', TobLigne.GetValue('GL_INDICEG'));
            TOBC1.PutValue('BCO_NUMORDRE', TobLigne.GetValue('GL_NUMORDRE'));
            TOBC1.PutValue('BCO_PHASETRA', NumPhase);
            TOBC1.PutValue('BCO_MOIS', Mois);
            TOBC1.PutValue('BCO_SEMAINE', Semaine);
            TOBC1.PutValue('BCO_DATEMOUV', DateMouv);
            TOBC1.PutValue('BCO_DPA',TOBC.GetValue('BCO_DPA'));
          end;
          TOBDC1.free;
        end;
        ferme (Q);
      end;
    end;
	FINALLY
  	TOBDETCONSO.free;
  END;
end;


procedure TGestionConso.MajoldConso;
var TOBC,TOBOC,TOBLC : TOB;
    LienPrec : double;
    req : String;
    QQ: Tquery;
    indice,Indice2 : integer;
begin
  TOBLC := TOB.Create ('LES ANCIENNES CONSO',nil,-1);
  TOBOldCOnso.ClearDetail;
  for indice := 0 to TOBConso.detail.count -1 do
  begin
    TOBC := TOBConso.detail[Indice];
    if TOBC.GetValue('BCO_INDICE') > 0 then continue;   // on ne traite pas les détails
    LienPrec := TOBC.GetValue('BCO_LIENTRANSFORME');
    if LienPrec <> 0 then
    begin
      TOBLC.ClearDetail;
      QQ := OpenSql ('SELECT * FROM CONSOMMATIONS WHERE BCO_NUMMOUV='+FloatToStr(LienPrec)+ ' AND BCO_INDICE=0',True);
      if not QQ.eof then
      begin
        TOBLC.LoadDetailDb ('CONSOMMATIONS','','',QQ,false);
        for Indice2 := 0 to TOBLC.detail.count -1 do
        begin
          TOBOC := TOB.Create ('CONSOMMATIONS',TobOldCOnso,-1);
          TOBOC.Dupliquer (TOBLC.detail[indice2],false,true);
          if TOBOC.GetValue('BCO_INDICE')=0 then
          begin
            TOBOC.PutValue('BCO_QUANTITE',TOBOC.GetValue('BCO_QUANTITE')-TOBC.GetVAlue('BCO_QUANTITE'));
            if TOBOC.getValue('BCO_QUANTITE') < 0 then TOBOC.PutValue('BCO_QUANTITE',0);
            if TOBOC.GetVAlue('BCO_QUANTITE') = 0 Then TOBOC.putValue('BCO_TRANSFORME','X');
          end;
          calculeLaLigne (TOBOC);
        end;
      end;
      ferme (QQ);
    end;
  end;
  if TOBOldConso.detail.count > 0 then
  begin
    if not TOBOLdConso.UpdateDB (false) then
    begin
      MessageValid := 'Erreur mise à jour CONSOS';
      V_PGI.IOError := OeUnknown;
    end;
  end;
  TOBLC.free;
end;

procedure TGestionConso.AffecteIndices;
var TOBC : TOB;
    i, Indice : integer;
    Nummouv : Double;
begin
  // tri de la tob avant exploration
  TOBConso.Detail.Sort('BCO_NUMMOUV;BCO_INDICE');
  // Réaffectation des indices
  Nummouv := -1;
  Indice := 0;
  for i := 0 to TOBConso.detail.count -1 do
  begin
    TOBC := TOBConso.detail[i];
    if TOBC <> nil then
    begin
      if TOBC.GetValue('BCO_NUMMOUV') <> Nummouv then
      begin
        Indice := TOBC.GetValue('BCO_INDICE');
        Nummouv := TOBC.GetValue('BCO_NUMMOUV');
      end else
      begin
        Inc(Indice);
        TOBC.PutValue('BCO_INDICE',Indice);
      end;
    end;
  end;
end;

procedure TGestionConso.MajTableConso;
Begin
  if TOBConso.detail.count = 0 then exit;
  if (GestionConso = TTcoDelete) or
     (gestionConso = TTcoModif) or
     (gestionConso = TTcoRegroup) or
     (GestionConso = TTcoTransform) then
  begin
    MajoldConso;
  end;

  if V_PGI.IOError = OeOK then
  begin
    // réaffectation indices
    AffecteIndices;
    TobConso.SetAllModifie(True);
    if not TOBConso.InsertDB (nil,true) then V_PGI.IOError := OeUnknown;
    if (GestionConso = TTcoTransform) or (GestionConso = TTcoModif) then
    begin
      ReajustePrixAchat;
    end;
  end;
End;

destructor TGestionConso.destroy;
begin
  inherited;
  TOBConso.free;
  TOBOldConso.free;
  TOBReceptionFou.free;
 	TOBReceptHorsLien.free;
end;

procedure TGestionConso.Clear;
begin
  TOBConso.ClearDetail ;
  TOBOldConso.ClearDetail;
  //
  TOBReceptionFou.cleardetail;
  TOBReceptHorsLien.clearDetail;
  //
end;

procedure TGestionConso.FlagRetourChantierAssocie (NumConsoRecep : double;LienVente : double);
var TOBC : TOB;
begin
	TOBC := TOBreceptionFou.findFirst(['_NUMRECEPFOUR_'],[NumConsoRecep],true);
  while TOBC <> nil do
  begin
    TOBC.putValue('BCO_LIENVENTE',LienVente);
    TOBC.PutValue('A TRAITER','X');
		TOBC := TOBreceptionFou.findNext(['_NUMRECEPFOUR_'],[NumConsoRecep],true);
  end;
end;

procedure TGestionConso.DefiniQteLivree(TOBL : TOB;QteLivre: double; LienVente : double);
var TOBE,TOBC : TOB;
    indice : integer;
    resteQte,QteLivrable,ResteQtePrec,QteReceptionne,QTeRetourne : double;
    RefPiece : string;
begin
  if QteLivre = 0 then exit;
  ResteQte := QteLivre;
  if TOBL.GetValue('GL_PIECEPRECEDENTE') = '' then
  begin
    TOBE := TOBreceptionFou.FindFirst (['REFLIVRCLI'],[EncodeRefPieceLivr (TOBL)],true);
  end else
  begin
  	refPiece := TOBL.GetValue('GL_PIECEPRECEDENTE');
    TOBE := TOBreceptionFou.FindFirst (['REFLIVRCLI'],[EncodeRefPieceLivr(RefPiece)],true);
  end;
  if TOBE = nil then exit;

  // --------------------------------------------------------------
  // Gestion des livraisons directes via receptions fournisseurs
  // --------------------------------------------------------------
  if (TOBL.GEtValue('GL_IDENTIFIANTWOL') = -1) then
  begin
  	// dans le cas d'une livraison sur chantier on va forcement livre la quantite receptionne --> d'ou
    TOBC := TOBE.FindFirst (['_REFLIVRFOU_'],[TOBL.GetValue('GL_PIECEORIGINE')],true);
    if TOBC <> nil then
    begin
      TOBE.PutValue ('A TRAITER','X');
      TOBC.PutValue ('BCO_QTELIVRE',ResteQte);
      TOBC.putValue('BCO_LIENVENTE',LienVente);
    	TOBC.PutValue('A TRAITER','X');
      PositionneEventuelDoublon (TOBC,resteQte);
      exit; // c finit pas la peine d'aller plus loin
    end;
  end;
  // --------------------------------------------------------------

  QteRetourne := 0;
  // gestion des retours fournisseur non rattaché a une reception
  TOBC := TOBE.FindFirst(['_NUMRECEPFOUR_'],[0],true);
  while TOBC <> nil do
  begin
    QteRetourne := QteRetourne + (TOBC.GetValue('BCO_QTERETOUR')*-1);
    TOBC.putValue('BCO_LIENVENTE',LienVente);
    TOBC.PutValue('A TRAITER','X');
  	TOBC := TOBE.Findnext(['_NUMRECEPFOUR_'],[0],true);
  end;

  ResteQte := ResteQte + QteRetourne; // on ajoute à la quantité a livrer les retours non rattachés

  for Indice := 0 to TOBE.detail.count -1 do
  begin
    TOBC := TOBE.detail[Indice];

    QteReceptionne := TOBC.getValue('BCO_QUANTITE') - TOBC.getValue('BCO_QTERETOUR');

    QteLivrable := QteReceptionne - TOBC.GEtValue('BCO_QTEVENTE');
    if QteLivrable <> 0 then
    begin
      TOBE.PutValue ('A TRAITER','X');
      resteQtePrec := ResteQte;
      ResteQte := ResteQte - QteLivrable;
      if ResteQte <= 0 then
      begin
      	TOBC.putValue ('BCO_QTELIVRE',TOBC.getValue('BCO_QTELIVRE')+ResteQtePrec);
        PositionneEventuelDoublon (TOBC,resteQtePrec);
      end else
      begin
      	TOBC.PutValue ('BCO_QTELIVRE',TOBC.getValue('BCO_QTELIVRE')+QteLivrable);
        PositionneEventuelDoublon (TOBC,QteLivrable);
      end;
      TOBC.putValue('BCO_LIENVENTE',LienVente);
    	TOBC.PutValue('A TRAITER','X');
      FlagRetourChantierAssocie (TOBC.GetValue('BCO_NUMMOUV'),LienVente);
      if resteQte <= 0 then break;
    end;
  end;
end;

procedure ReajusteRecptFromLivr (TOBL : TOB);
var TOBRetDet,TOBRet,TOBRD,TOBR,TOBC : TOB;
    indice : integer;
    Req : string;
    QQ : Tquery;
begin
  TOBC := TOB.Create ('CONSOMMATIONS',nil,-1);
  req := 'SELECT * FROM CONSOMMATIONS WHERE BCO_NUMMOUV='+
          FloatToStr(TOBL.GetValue('BLP_NUMMOUV'));
  QQ := OpenSql (Req,True);
  if not QQ.eof then
  begin
     TOBC.SelectDB ('',QQ);
  end else
  begin
    TOBC.free;
    Ferme (QQ);
    Exit;
  end;
  Ferme (QQ);

  TOBRetDet := TOB.Create ('LES RECEPTIONS DETAILLES',nil,-1);
  TOBRet := TOB.Create ('LES RECEPTIONS',nil,-1);
  req := 'SELECT * FROM CONSOMMATIONS WHERE BCO_LIENVENTE='+
          FloatToStr(TOBC.GetValue('BCO_NUMMOUV'))+
          ' ORDER BY BCO_NUMMOUV,BCO_INDICE';
  QQ:= OpenSql (Req ,True);
  TRY
    if not QQ.eof then
    begin
      TOBRetDet.LoadDetailDB ('CONSOMMATIONS','','',QQ,false);
      Ferme (QQ);
      for Indice := 0 to TOBRetDet.detail.count -1 do
      begin
        TOBRD := TOBRetDet.detail[Indice];
        TOBR := TOB.create ('CONSOMMATIONS',TOBRET,-1);
        req := 'SELECT * FROM CONSOMMATIONS WHERE BCO_NUMMOUV='+
               FloatToStr(TOBRD.GetValue('BCO_NUMMOUV'))+
               ' AND BCO_INDICE=0';
        QQ := OpenSql (req,True);
        TOBR.SelectDB ('',QQ);
        TOBR.PutValue('BCO_TRAITEVENTE','-');
        TOBR.PutValue('BCO_QTEVENTE',TOBR.GetValue('BCO_QTEVENTE') - TOBRD.GetValue('BCO_QTEVENTE'));
      end;
      if not TOBRet.UpdateDB then BEGIN V_PGI.IOError := OeUnknown; Exit; END;
      if not TOBRetDet.DeleteDB then BEGIN V_PGI.IOError := OeUnknown; Exit; END;
    end;
  FINALLY
   Ferme (QQ);
   TOBRetDet.free;
   TOBRet.free;
   TOBC.free;
  END;
end;

procedure TGestionConso.ReajusteReception;
var indice,Indice2 : integer;
    TOBC,TOBNC,TOBANC,TOBE : TOB;
    TOBDetReception : TOB;
    QQ : TQuery;
    Req : String;
    LastNum : double;
begin
  TOBDetReception := TOB.Create ('LE DETAIL DE LA RECEPTION',nil,-1);
  TRY
    if TOBreceptionFou.detail.count = 0 then exit;
    for Indice := 0 to TOBreceptionFou.detail.count -1 do
    begin
      TOBE := TOBreceptionFou.detail[Indice];
      if TOBE.GetValue('A TRAITER') = 'X' then
      begin
        for Indice2 := 0 to TOBE.detail.count -1 do
        BEGIN
          TOBC := TOBE.detail[Indice2];

          //if (TOBC.GetValue('BCO_QTELIVRE') = 0) and (TOBC.GetValue('_NATURE_')='RECEPTION')  then continue;
          if TOBC.GetValue('A TRAITER') <>'X' then continue;

          Req := 'SELECT * FROM CONSOMMATIONS WHERE BCO_NUMMOUV='+
                  FloatToStr(TOBC.GetValue('BCO_NUMMOUV')) +
                  ' ORDER BY BCO_NUMMOUV,BCO_INDICE';
          QQ := OpenSql (Req,true);
          if not QQ.eof then
          begin
            // Chargement du detail de la reception par livraison client
            TOBDetReception.LoadDetailDB ('CONSOMMATIONS','','',QQ,false);
            lastNum := TOBDetReception.detail[TOBDetReception.detail.count -1].GetValue('BCO_INDICE')+1;
            TOBANC := TOBDetReception.detail[0];
            // Création d'une nouvelle ligne contenant le detail de livraison pour cette reception ou facture..
            TOBNC := TOB.Create ('CONSOMMATIONS',TOBCONSO,-1);
            TOBNC.dupliquer (TOBANC,false,true);
            TOBNC.PutValue('BCO_INDICE',lastNum);
            TOBNC.PutValue('BCO_QUANTITE',TOBC.GetValue('BCO_QTELIVRE'));
            TOBNC.PutValue('BCO_QTEVENTE',TOBC.GetValue('BCO_QTELIVRE'));
            TOBNC.PutValue('BCO_LIENVENTE',TOBC.GetValue('BCO_LIENVENTE'));
            TOBNC.PutValue('BCO_TRAITEVENTE','X');
            TOBNC.PutValue('BCO_TRANSFORME','X');
            calculeLaLigne (TOBNC);

            // Réajustement de la réception principale
            TOBANC := TOBDetReception.detail[0]; // l'indice = 0 se trouve forcement a l'indice 0..logique non ?
            TOBANC.putValue('BCO_QTEVENTE',TOBANC.GetValue('BCO_QTEVENTE')+TOBC.getValue('BCO_QTELIVRE'));
            if TOBANC.getValue('BCO_QTEVENTE')>=(TOBANC.GetValue('BCO_QUANTITE')-TOBC.GetValue('BCO_QTERETOUR')) then
            		TOBANC.putValue('BCO_TRAITEVENTE','X');
            // mise à jour
            if not TOBANC.UpdateDB (false) then BEGIN V_PGI.IOError:=OeUnknown; Break; End;
          end;
        END;
      END;
    end;
  FINALLY
    TOBDetReception.Free;
  END;
end;

procedure TGestionConso.MemoriseLienReception(TOBL: TOB; TOBReception : TOB);
var TOBE,TOBC : TOB;
		Naturepiece : string;
    NumIndice : integer;
    QteRecep,NumMouvRecep,QteDejaLivr : double;
    Numrecep : double;
    refPiece : string;
begin
	NumMouvRecep := TOBReception.GetValue('BCO_NUMMOUV');
	NumIndice := TOBReception.GetValue('BCO_INDICE');
	QteRecep := TOBReception.GetValue('BCO_QUANTITE');
	QteDejaLivr := TOBReception.GetValue('BCO_QTEVENTE');
  if QteRecep = 0 then exit;

  Naturepiece := TOBReception.GetValue('GL_NATUREPIECEG');

  if TOBL.GetValue('GL_PIECEPRECEDENTE') = '' then
  begin
    TOBE := TOBreceptionFou.FindFirst (['REFLIVRCLI'],[EncodeRefPieceLivr (TOBL)],true);
  end else
  begin
  	refPiece := TOBL.GetValue('GL_PIECEPRECEDENTE');
    TOBE := TOBreceptionFou.FindFirst (['REFLIVRCLI'],[EncodeRefPieceLivr(RefPiece)],true);
  end;

  if (TOBE = nil) then
  begin
    TOBE := TOB.create ('UNE LIGNE LIVRAISON CLIENT',TOBReceptionFou,-1);

    if TOBL.GetValue('GL_PIECEPRECEDENTE') = '' then
    begin
      TOBE.AddChampSupValeur ('REFLIVRCLI',EncodeRefPieceLivr (TOBL));
    end else
    begin
    	refPiece := TOBL.GetValue('GL_PIECEPRECEDENTE');
      TOBE.AddChampSupValeur ('REFLIVRCLI',EncodeRefPieceLivr(RefPiece));
    end;
    TOBE.AddChampSupValeur ('A TRAITER','-');
  end;

  if Naturepiece <> 'BFA' then
  begin
  	// cas des receptions fournisseurs
    TOBC := TOB.Create ('LES CONSOS RECEP FOUR',TOBE,-1);
    TOBC.AddChampSupValeur('_REFLIVRFOU_',EncodeRefPiece (TOBReception,0,true));
    TOBC.AddChampSupValeur('_NUMRECEPFOUR_',-1);
    TOBC.AddChampSupValeur('BCO_NUMMOUV',NumMouvrecep);
    TOBC.AddChampSupValeur('BCO_INDICE',NumIndice);
    TOBC.AddChampSupValeur('BCO_QUANTITE',QteRecep);
    TOBC.AddChampSupValeur('BCO_QTEVENTE',QteDejalivr);
    TOBC.AddChampSupValeur('BCO_QTERETOUR',0);
    TOBC.AddChampSupValeur('BCO_QTELIVRE',0);
    TOBC.AddChampSupValeur('BCO_LIENVENTE',0);
    TOBC.AddChampSupValeur('_NATURE_','RECEPTION');
    TOBC.AddChampSupValeur('A TRAITER','-');
    TOBC.AddChampSupValeur('USABLE','X');
  end else
  begin
  	// cas du retour fournisseur
    TOBC := TOBE.FindFirst(['_REFLIVRFOU_'],[TOBReception.GetValue('GL_PIECEPRECEDENTE')],True);
    if TOBC <> nil then
    begin
    	// mise a jour de la quantite retourne sur la reception
    	TOBC.PutValue('BCO_QTERETOUR',TOBC.GetValue('BCO_QTERETOUR')+(QteRecep*-1));
      Numrecep := TOBC.GetValue('BCO_NUMMOUV');
      // Creation de la ligne retour
      TOBC := TOB.Create ('LES CONSOS RECEP FOUR',TOBE,-1);
    	TOBC.AddChampSupValeur('_REFLIVRFOU_','');
    	TOBC.AddChampSupValeur('_NUMRECEPFOUR_',Numrecep);
      TOBC.AddChampSupValeur('BCO_NUMMOUV',NumMouvRecep);
    	TOBC.AddChampSupValeur('BCO_INDICE',NumIndice);
      TOBC.AddChampSupValeur('BCO_QUANTITE',0);
      TOBC.AddChampSupValeur('BCO_QTEVENTE',0);
      TOBC.AddChampSupValeur('BCO_QTERETOUR',(QteRecep*-1));
      TOBC.AddChampSupValeur('BCO_QTELIVRE',0);
      TOBC.AddChampSupValeur('BCO_LIENVENTE',0);
      TOBC.AddChampSupValeur('_NATURE_','RETOUR');
    	TOBC.AddChampSupValeur('A TRAITER','-');
    	TOBC.AddChampSupValeur('USABLE','X');
    end else
    begin
      TOBC := TOB.Create ('LES CONSOS RECEP FOUR',TOBE,-1);
    	TOBC.AddChampSupValeur('_REFLIVRFOU_','');
    	TOBC.AddChampSupValeur('_NUMRECEPFOUR_',0);
      TOBC.AddChampSupValeur('BCO_NUMMOUV',NumMouvRecep);
    	TOBC.AddChampSupValeur('BCO_INDICE',NumIndice);
      TOBC.AddChampSupValeur('BCO_QUANTITE',0);
      TOBC.AddChampSupValeur('BCO_QTEVENTE',0);
      TOBC.AddChampSupValeur('BCO_QTERETOUR',(QteRecep*-1));
      TOBC.AddChampSupValeur('BCO_QTELIVRE',0);
      TOBC.AddChampSupValeur('BCO_LIENVENTE',0);
      TOBC.AddChampSupValeur('_NATURE_','RETOUR');
    	TOBC.AddChampSupValeur('A TRAITER','-');
    	TOBC.AddChampSupValeur('USABLE','X');
    end;
  end;
end;

procedure TGestionConso.InitReceptions;
begin
  TOBReceptionFou.ClearDetail;
end;

procedure TGestionConso.recupereReceptions(TOBL: TOB);
var Requete : string;
    QQ : Tquery;
    TOBDetRec,TOBDR,TOBE,TOBC : TOB;
    RefOrigine : string;
    Indice : integer;
    Numrecep : double;
begin
	if not IsLivraisonClient (TOBL) then exit;
  RefOrigine := TOBL.GetValue('GL_PIECEPRECEDENTE');
//  if RefOrigine = '' then  RefOrigine := EncodeRefPiece (TOBL,0,True);
//  RefOrigine := TOBL.GetValue('GL_PIECEORIGINE'); if RefOrigine = '' then exit;
//  TOBE := TOBreceptionFou.FindFirst (['REFLIVRCLI'],[EncodeRefPiece (TOBL,0,True)],true);

  if RefOrigine = '' then
  begin
  	TOBE := TOBreceptionFou.FindFirst (['REFLIVRCLI'],[EncodeRefPieceLivr (TOBL)],true);
  end else
  begin
  	TOBE := TOBreceptionFou.FindFirst (['REFLIVRCLI'],[EncodeRefPieceLivr(RefOrigine)],true);
  end;
  if TOBE = nil then
  begin
    TOBE := TOB.create ('UNE LIGNE LIVRAISON CLIENT',TOBReceptionFou,-1);
//    TOBE.AddChampSupValeur ('REFLIVRCLI',EncodeRefPiece (TOBL,0,true));
      if RefOrigine = '' then
      begin
        TOBE.AddChampSupValeur ('REFLIVRCLI',EncodeRefPieceLivr (TOBL),false);
      end else
      begin
        TOBE.AddChampSupValeur ('REFLIVRCLI',EncodeRefPieceLivr(RefOrigine),false);
      end;
    TOBE.AddChampSupValeur ('A TRAITER','-');
  end;

  TOBDetRec := TOB.Create ('LES LIGNES',nil,-1);
  requete := 'SELECT GL_QTERESTE AS QTERECEP'+
             ',(GL_QTESTOCK*GL_PUHTNET) AS MTACHAT, GL_QUALIFQTEACH,GL_NATUREPIECEG,GL_DATEPIECE,'+
             'GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE,GL_NUMORDRE,GL_PIECEPRECEDENTE,'+
             'BCO_NUMMOUV,BCO_INDICE,BCO_QUANTITE,BCO_QTEVENTE,BCO_TRANSFORME,BCO_TRAITEVENTE,BCO_LIENVENTE '+
             'FROM LIGNE '+
             'LEFT JOIN CONSOMMATIONS ON '+
             '(BCO_NATUREPIECEG=GL_NATUREPIECEG) AND (BCO_SOUCHE=GL_SOUCHE) AND (BCO_NUMERO=GL_NUMERO) '+
             'AND (BCO_NUMORDRE=GL_NUMORDRE) '+
             'WHERE ';
  if refOrigine <> '' then
  begin
  	requete := requete + 'GL_PIECEORIGINE="'+RefOrigine+ '" AND '
  end else
  begin
  	requete := requete + 'GL_AFFAIRE="'+TOBL.GetValue('GL_AFFAIRE')+
    					'" AND GL_ARTICLE="'+TOBL.GetValue('GL_ARTICLE')+'" AND ';
  end;
  Requete := requete +'((BCO_LIENVENTE='+InttoStr(TOBL.GetValue('BLP_NUMMOUV'))+
  					') OR (BCO_LIENVENTE=0 '+(*AND BCO_QUANTITE > BCO_QTEVENTE *) ' AND BCO_TRANSFORME <> "X")) '+
             'AND GL_NATUREPIECEG IN ('+GetPieceAchat(True,False)+') '+
             (*'AND GL_VIVANTE="X" AND BCO_TRANSFORME<>"X" *)
             'AND GL_DATEPIECE <= "'+USDateTime(TOBL.GetValue('GL_DATEPIECE'))+'" ORDER BY GL_DATEPIECE';
             // les lignes de receptions avec indice à 0
  QQ := OpenSql (requete,true);
  TRY
    if not QQ.eof then
    begin
      // recupère le detail des réceptions de la ligne de document initiale
      TOBDetRec.loadDetailDb ('LIGNE','','',QQ,false);
      for Indice := 0 to TOBDetRec.detail.count -1 do
      begin
        TOBDR := TOBDetRec.detail[Indice];
        if TOBDR.GetValue('GL_NATUREPIECEG') <> 'BFA' Then
        begin
          TOBC := TOB.Create ('LES CONSOS RECEP FOUR',TOBE,-1);

          TOBC.AddChampSupValeur('_REFLIVRFOU_',EncodeRefPiece(TOBDR)); // reference de la recption en cours
          TOBC.AddChampSupValeur('_NUMRECEPFOUR_',-1);                    // reference de la reception pour le retour
          TOBC.AddChampSupValeur('BCO_NUMMOUV',TOBDR.GetValue('BCO_NUMMOUV'));
          TOBC.AddChampSupValeur('BCO_INDICE',TOBDR.GetValue('BCO_INDICE'));
          TOBC.AddChampSupValeur('BCO_QUANTITE',TOBDR.GetValue('BCO_QUANTITE'));
          TOBC.AddChampSupValeur('BCO_QTEVENTE',TOBDR.GetValue('BCO_QTEVENTE'));
          TOBC.AddChampSupValeur('BCO_QTELIVRE',0);
          TOBC.AddChampSupValeur('BCO_QTERETOUR',0);
          TOBC.AddChampSupValeur('BCO_LIENVENTE',TOBL.GetValue('BLP_NUMMOUV'));
          TOBC.AddChampSupValeur('_NATURE_','RECEPTION');
    			TOBC.AddChampSupValeur('A TRAITER','-');
    			TOBC.AddChampSupValeur('USABLE','X');
        end else
        begin
        	TOBC := TOBE.findFirst(['_REFLIVRFOU_'],[TOBDR.GetValue('GL_PIECEPRECEDENTE')],true);
          if TOBC <> nil then
          begin
          	TOBC.PutValue('BCO_QTERETOUR',TOBC.GetValue('BCO_QTERETOUR')+(TOBDR.GetValue('BCO_QUANTITE')*-1));
            Numrecep := TOBC.GetValue('BCO_NUMMOUV');
            // Creation de la ligne retour
            TOBC := TOB.Create ('LES CONSOS RECEP FOUR',TOBE,-1);
            TOBC.AddChampSupValeur('_REFLIVRFOU_','');
            TOBC.AddChampSupValeur('_NUMRECEPFOUR_',Numrecep);
          	TOBC.AddChampSupValeur('BCO_INDICE',TOBDR.GetValue('BCO_INDICE'));
            TOBC.AddChampSupValeur('BCO_NUMMOUV',TOBDR.GetValue('BCO_NUMMOUV'));
            TOBC.AddChampSupValeur('BCO_QUANTITE',0);
            TOBC.AddChampSupValeur('BCO_QTEVENTE',0);
            TOBC.AddChampSupValeur('BCO_QTERETOUR',(TOBDR.GetValue('BCO_QUANTITE')*-1));
            TOBC.AddChampSupValeur('BCO_QTELIVRE',0);
            TOBC.AddChampSupValeur('BCO_LIENVENTE',0);
            TOBC.AddChampSupValeur('_NATURE_','RETOUR');
    				TOBC.AddChampSupValeur('A TRAITER','-');
    				TOBC.AddChampSupValeur('USABLE','X');
          end else
          begin
            TOBC := TOB.Create ('LES CONSOS RECEP FOUR',TOBE,-1);
            TOBC.AddChampSupValeur('_REFLIVRFOU_','');
            TOBC.AddChampSupValeur('_NUMRECEPFOUR_',0);
            TOBC.AddChampSupValeur('BCO_NUMMOUV',TOBDR.GetValue('BCO_NUMMOUV'));
          	TOBC.AddChampSupValeur('BCO_INDICE',TOBDR.GetValue('BCO_INDICE'));
            TOBC.AddChampSupValeur('BCO_QUANTITE',0);
            TOBC.AddChampSupValeur('BCO_QTEVENTE',0);
            TOBC.AddChampSupValeur('BCO_QTERETOUR',(TOBDR.GetValue('BCO_QUANTITE')*-1));
            TOBC.AddChampSupValeur('BCO_QTELIVRE',0);
            TOBC.AddChampSupValeur('BCO_LIENVENTE',0);
            TOBC.AddChampSupValeur('_NATURE_','RETOUR');
    				TOBC.AddChampSupValeur('A TRAITER','-');
    				TOBC.AddChampSupValeur('USABLE','X');
          end;
        end;
      end;

      Ferme (QQ);
      TOBDetRec.cleardetail;
      if TOBL.GetValue('BLP_NUMMOUV') > 0 then
      begin
        requete := 'SELECT * FROM CONSOMMATIONS WHERE BCO_LIENVENTE='+
                FloatToStr(TOBL.GetValue('BLP_NUMMOUV'));
        QQ := OpenSql (requete,true);
        if not QQ.eof then
        begin
          TOBDetRec.LoadDetailDB ('CONSOMMATIONS','','',QQ,false);
          for Indice := 0 to TOBDetRec.detail.count -1 do
          begin
            TOBDR := TOBDetRec.detail[Indice]; // ligne de consos liés a la réceptions et la livraison
            TOBC := TOBE.FindFirst(['BCO_NUMMOUV'],[TOBDR.GEtVAlue('BCO_NUMMOUV')],True);
            if TOBC <> nil then
            begin
              TOBC.putValue ('BCO_QTEVENTE',TOBC.getValue('BCO_QTEVENTE')-TOBDR.GetValue('BCO_QTEVENTE'));
            end;
          end;
        end;
      end;
    end;
  FINALLY
    Ferme (QQ);
    TOBDetRec.free;
  END;
end;

function 	TGestionConso.IsLivraisonClient (TOBL : TOB) : boolean;
begin
	result := false;
  if TOBL.NomTable = 'LIGNE' then
  begin
  	if Pos(TOBL.GetValue('GL_NATUREPIECEG'),GetNaturePieceLBT(false)) > 0 then result := true;
  end else if TOBL.nomTable = 'PIECE' then
  begin
  	if Pos(TOBL.GetValue('GP_NATUREPIECEG'),GetNaturePieceLBT(false)) > 0 then result := true;
  end;
end;

procedure TGestionConso.ReajustePrixAchat;
var Indice : integer;
    TOBC,TOBRC,TOBLRC : TOB;
    QQ : Tquery;
    CoefPaPr : double;
    Req : string;
begin
  TOBLRC := TOB.Create ('LES LIVRAISONS',nil,-1);
  TRY
    for Indice := 0 to TOBConso.detail.count -1 do
    begin
      TOBC := TOBConso.detail[Indice];
      if (TOBC.GetValue('BCO_NATUREPIECEG')='FF') and (TOBC.GetValue('BCO_LIENVENTE')> 0) then
      begin
        Req := 'SELECT * FROM CONSOMMATIONS WHERE BCO_NUMMOUV='+
                FloatToStr(TOBC.GetValue('BCO_LIENVENTE')) +
               ' AND BCO_INDICE = 0';
        QQ := OpenSql (req,True);
        if not QQ.eof then
        begin
          TOBRC := TOB.Create ('CONSOMMATIONS',TOBLRC,-1);
          TOBRC.SelectDB ('',QQ);
          if TOBRC.GetVAlue('BCO_DPA') = 0 then
          begin
          	CoefPaPr := 1;
          end else
          begin
          	CoefPaPR := TOBRC.GetVAlue('BCO_DPR')/TOBRC.GetVAlue('BCO_DPA');
          end;
          TOBRC.putValue('BCO_DPA',TOBC.GetValue('BCO_DPA'));
          TOBRC.putValue('BCO_DPR',Arrondi(TOBC.GetValue('BCO_DPA')*CoefPaPr,V_PGI.OkDecP));
          calculeLaLigne(TOBRC);
        end;
      end;
    end;
    if TOBLRC.detail.count > 0 then
    if not TOBLRC.UpdateDB Then V_PGI.IoError := OeUnknown;
  FINALLY
    TOBLRC.Free;
  END;
end;

procedure TGestionConso.recupereReceptionsHorsLien (TOBPiece : TOB;Laffaire : String; Larticle : string='');
var Requete : string;
    QQ : Tquery;
    TOBDetRec,TOBDR,TOBC : TOB;
    RefOrigine : string;
    Indice : integer;
    Numrecep : double;
    cledoc : R_CLEDOC;
begin
  if TOBPiece = nil then exit;
	if not IsLivraisonClient (TOBPiece) then exit;
  TOBDetRec := TOB.Create ('LES LIGNES',nil,-1);
  requete := 'SELECT LIG.*,(GL_QTESTOCK*GL_PUHTNET) AS MTACHAT,'+
             'BCO_NUMMOUV,BCO_INDICE,BCO_QUANTITE,BCO_QTEVENTE,BCO_TRANSFORME,BCO_TRAITEVENTE '+
             'FROM CONSOMMATIONS '+
             'LEFT JOIN LIGNE LIG ON '+
             '(GL_NATUREPIECEG=BCO_NATUREPIECEG) AND (GL_SOUCHE=BCO_SOUCHE) AND (GL_NUMERO=BCO_NUMERO) '+
             'AND (GL_NUMORDRE=BCO_NUMORDRE) '+
             'WHERE BCO_AFFAIRE="'+Laffaire+ '" AND BCO_TRAITEVENTE<>"X"';
  if Larticle <> '' then
  begin
  	 requete := requete + ' AND GL_ARTICLE="'+Larticle+'" ';
  end;
   requete := requete + (*'AND BCO_QUANTITE >= BCO_QTEVENTE *)' AND BCO_TRANSFORME<>"X" '+
  										 'AND GL_NATUREPIECEG IN ('+GetPieceAchat(True,False,true,true)+') '+
                       'ORDER BY GL_ARTICLE,GL_DATEPIECE,GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_NUMORDRE';
             // les lignes de receptions avec indice à 0
  QQ := OpenSql (requete,true);
  TRY
    if not QQ.eof then
    begin
  {
    	if First then
      begin
        TOBE := TOB.create ('LES LIGNES HORS LIENS VTE',TOBReceptHorsLien,-1);
        first := false;
      end;
  }
  // recupère le detail des réceptions non liés à des besoins de chantiers
      TOBDetRec.loadDetailDb ('LIGNE','','',QQ,false);
      for Indice := 0 to TOBDetRec.detail.count -1 do
      begin
        TOBDR := TOBDetRec.detail[Indice];
        RefOrigine := TOBDR.GetValue('GL_PIECEORIGINE');
        if RefOrigine <> '' then
        begin
        	// on passe les réceptions en provenance d'un besoin de chantier
          DecodeRefPiece (RefOrigine,Cledoc);
          if Pos(Cledoc.NaturePiece,GetPieceAchat (true,True,false,true,true)) = 0 then continue;
        end;
        if TOBDR.GetValue('GL_NATUREPIECEG') <> 'BFA' Then
        begin
          TOBC := TOB.Create ('LIGNE',TOBReceptHorsLien,-1);
          TOBC.dupliquer (TOBDR,false,true);
          TOBC.AddChampSupValeur('_REFLIVRFOU_',EncodeRefPiece(TOBDR)); // reference de la recption en cours
          TOBC.AddChampSupValeur('_LIBREFLIVRFOU_',LibellePiece(TOBDR)); // reference de la recption en cours
          TOBC.AddChampSupValeur('_NUMRECEPFOUR_',-1);                    // reference de la reception pour le retour
          TOBC.AddChampSupValeur('BCO_NUMMOUV',TOBDR.GetValue('BCO_NUMMOUV'));
          TOBC.AddChampSupValeur('BCO_INDICE',TOBDR.GetValue('BCO_INDICE'));
          TOBC.AddChampSupValeur('BCO_QUANTITE',TOBDR.GetValue('BCO_QUANTITE'));
          TOBC.AddChampSupValeur('BCO_QTEVENTE',TOBDR.GetValue('BCO_QTEVENTE'));
          TOBC.AddChampSupValeur('BCO_QTELIVRE',0);
          TOBC.AddChampSupValeur('BCO_QTERETOUR',0);
          TOBC.AddChampSupValeur('BCO_LIENVENTE',0);
          TOBC.AddChampSupValeur('MTACHAT',TOBDR.GetValue('MTACHAT'));
          TOBC.AddChampSupValeur('_NATURE_','RECEPTION');
    			TOBC.AddChampSupValeur('A TRAITER','-');
    			TOBC.AddChampSupValeur('USABLE','X');
        end else
        begin
        	TOBC := TOBReceptHorsLien.findFirst(['_REFLIVRFOU_'],[TOBDR.GetValue('GL_PIECEPRECEDENTE')],true);
          if TOBC <> nil then
          begin
          	TOBC.PutValue('BCO_QTERETOUR',TOBC.GetValue('BCO_QTERETOUR')+(TOBDR.GetValue('BCO_QUANTITE')*-1));
            Numrecep := TOBC.GetValue('BCO_NUMMOUV');
            // Creation de la ligne retour
            TOBC := TOB.Create ('LIGNE',TOBReceptHorsLien,-1);
            TOBC.dupliquer (TOBDR,false,true);
            TOBC.AddChampSupValeur('_REFLIVRFOU_','');
          	TOBC.AddChampSupValeur('_LIBREFLIVRFOU_',''); // reference de la recption en cours
            TOBC.AddChampSupValeur('_NUMRECEPFOUR_',Numrecep);
            TOBC.AddChampSupValeur('BCO_NUMMOUV',TOBDR.GetValue('BCO_NUMMOUV'));
          	TOBC.AddChampSupValeur('BCO_INDICE',TOBDR.GetValue('BCO_INDICE'));
            TOBC.AddChampSupValeur('BCO_QUANTITE',0);
            TOBC.AddChampSupValeur('BCO_QTEVENTE',0);
            TOBC.AddChampSupValeur('BCO_QTERETOUR',(TOBDR.GetValue('BCO_QUANTITE')*-1));
            TOBC.AddChampSupValeur('BCO_QTELIVRE',0);
            TOBC.AddChampSupValeur('BCO_LIENVENTE',0);
          	TOBC.AddChampSupValeur('MTACHAT',TOBDR.GetValue('MTACHAT') * -1);
            TOBC.AddChampSupValeur('_NATURE_','RETOUR');
    				TOBC.AddChampSupValeur('A TRAITER','-');
    				TOBC.AddChampSupValeur('USABLE','X');
          end else
          begin
            TOBC := TOB.Create ('LIGNE',TOBReceptHorsLien,-1);
            TOBC.dupliquer (TOBDR,false,true);
            TOBC.AddChampSupValeur('_REFLIVRFOU_','');
          	TOBC.AddChampSupValeur('_LIBREFLIVRFOU_',''); // reference de la recption en cours
            TOBC.AddChampSupValeur('_NUMRECEPFOUR_',0);
            TOBC.AddChampSupValeur('BCO_NUMMOUV',TOBDR.GetValue('BCO_NUMMOUV'));
          	TOBC.AddChampSupValeur('BCO_INDICE',TOBDR.GetValue('BCO_INDICE'));
            TOBC.AddChampSupValeur('BCO_QUANTITE',0);
            TOBC.AddChampSupValeur('BCO_QTEVENTE',0);
            TOBC.AddChampSupValeur('BCO_QTERETOUR',(TOBDR.GetValue('BCO_QUANTITE')*-1));
          	TOBC.AddChampSupValeur('MTACHAT',TOBDR.GetValue('MTACHAT') * -1);
            TOBC.AddChampSupValeur('BCO_QTELIVRE',0);
            TOBC.AddChampSupValeur('BCO_LIENVENTE',0);
            TOBC.AddChampSupValeur('_NATURE_','RETOUR');
    				TOBC.AddChampSupValeur('A TRAITER','-');
    				TOBC.AddChampSupValeur('USABLE','X');
          end;
        end;
      end;
      Ferme (QQ);
      TOBDetRec.cleardetail;
    end;
  FINALLY
    Ferme (QQ);
    TOBDetRec.free;
  END;
end;

function TGestionConso.LibellePiece(TOBDR : TOB) : string;
begin
	result := GetInfoParPiece (TOBDR.GetValue('GL_NATUREPIECEG'),'GPP_LIBELLE')+' N° '+InttoStr(TOBDR.getValue('GL_NUMERO'))+' du '+DateToStr(TOBDR.getValue('GL_DATEPIECE'));
end;


procedure TGestionConso.recupereReceptionsHorsLien (TOBPiece,TOBAffaire : TOB);
var Laffaire : string;
begin
	TOBReceptHorsLien.ClearDetail;
  Laffaire := TOBAffaire.getValue('AFF_AFFAIRE');
  recupereReceptionsHorsLien (TOBpiece,Laffaire);
end;

procedure TGestionConso.DefiniReceptionsFromHlienAssocie(TOBL,TOBDetRec : TOB);
var
    TOBDR,TOBE,TOBC : TOB;
    Indice : integer;
begin
	if not IsLivraisonClient (TOBL) then exit;
  TOBE := TOBreceptionFou.FindFirst (['REFLIVRCLI'],[EncodeRefPieceLivr (TOBL)],true);
  if TOBE = nil then
  begin
    TOBE := TOB.create ('UNE LIGNE LIVRAISON CLIENT',TOBReceptionFou,-1);
    TOBE.AddChampSupValeur ('REFLIVRCLI',EncodeRefPieceLivr (TOBL),false);
    TOBE.AddChampSupValeur ('A TRAITER','X');
  end;

  for Indice := 0 to TOBDetRec.detail.count -1 do
  begin
    TOBDR := TOBDetRec.detail[Indice];
    TOBC := TOB.Create ('LES CONSOS RECEP FOUR',TOBE,-1);

    TOBC.AddChampSupValeur('_REFLIVRFOU_',TOBDR.GetValue('_REFLIVRFOU_')); // reference de la recption en cours
    TOBC.AddChampSupValeur('_NUMRECEPFOUR_',TOBDR.GetValue('_NUMRECEPFOUR_')); // reference de la reception pour le retour
    TOBC.AddChampSupValeur('BCO_NUMMOUV',TOBDR.GetValue('BCO_NUMMOUV'));
    TOBC.AddChampSupValeur('BCO_INDICE',TOBDR.GetValue('BCO_INDICE'));
    TOBC.AddChampSupValeur('BCO_QUANTITE',TOBDR.GetValue('BCO_QUANTITE'));
    TOBC.AddChampSupValeur('BCO_QTEVENTE',TOBDR.GetValue('BCO_QTEVENTE'));
    TOBC.AddChampSupValeur('BCO_QTELIVRE',TOBDR.GetValue('BCO_QTELIVRE'));
    TOBC.AddChampSupValeur('BCO_QTERETOUR',TOBDR.GetValue('BCO_QTERETOUR'));
    TOBC.AddChampSupValeur('BCO_LIENVENTE',TOBDR.GetValue('BCO_LIENVENTE'));
    TOBC.AddChampSupValeur('_NATURE_',TOBDR.GetValue('_NATURE_'));
    TOBC.AddChampSupValeur('A TRAITER','-');
    TOBC.AddChampSupValeur('USABLE','X');
  end;
end;


function TGestionConso.NbrReceptionsHorsLien: integer;
begin
	result := TOBReceptHorsLien.detail.count;
end;

function TGestionConso.GetTOBrecepHLien : TOB;
begin
	result := TOBReceptHorsLien;
end;

procedure TGestionConso.RepertorieReceptionsHLiensFromLIV (TOBPiece : TOB ; Affaire : String);
var Indice : integer;
		Larticle : string;
    TOBL : TOB;
begin
	for Indice := 0 to TOBPiece.detail.count -1 do
  begin
  	TOBL := TOBPiece.detail[Indice];
    larticle := TOBL.getValue('GL_ARTICLE');
    if larticle = '' then continue;
    recupereReceptionsHorsLien (TOBL,Affaire,Larticle);
    DefiniReceptionsFromHlienAssocie (TOBL,TOBReceptHorsLien);
    TOBReceptHorsLien.clearDetail;
  end;
end;

function TGestionConso.GetQteLivrable (TOBL : TOB; WithGeneration : boolean= false) : double;
var TOBE,TOBC : TOB;
		indice : integer;
    Affaire :string ;
begin
	if WithGeneration then
  begin
  	Affaire := TOBL.GetValue('GL_AFFAIRE');
		RepertorieReceptionsHLiensFromLIV (TOBL,Affaire);
	end;
	result := 0;
  TOBE := TOBreceptionFou.FindFirst (['REFLIVRCLI'],[EncodeRefPieceLivr (TOBL)],true);
  if TOBE = nil then exit;
  for Indice := 0 to TOBE.detail.count -1 do
  begin
  	TOBC := TOBE.detail[Indice];
    if TOBC.GetValue('BCO_INDICE')= 0 then
    begin
      if TOBC.GetValue('_REFLIVRFOU_')='' then
      begin
        result := result - TOBC.GetValue('BCO_QTERETOUR');
      end else
      begin
        result := result + TOBC.GetValue('BCO_QUANTITE')
                         - TOBC.GetValue('BCO_QTEVENTE');
      end;
    end;
  end;
end;

procedure TGestionConso.PositionneEventuelDoublon (TOBCourante : TOB;Qte : double);
var NumMouv : double;
		Indice : integer;
    TOBFC : TOB;
begin
	NumMouv := TOBCOurante.GetValue('BCO_NUMMOUV');
  Indice := TOBCOurante.GetValue('BCO_INDICE');
	TOBFC := TOBReceptionFou.findfirst(['BCO_NUMMOUV','BCO_INDICE'],[NumMouv,Indice],True);
  if TOBFC = nil then exit;
  repeat
  	if TOBFC <> TOBCourante then
    begin
    	TOBFC.PutValue('BCO_QUANTITE',TOBFC.GETValue('BCO_QUANTITE')-Qte);
    end;
		TOBFC := TOBReceptionFou.findNext(['BCO_NUMMOUV','BCO_INDICE'],[NumMouv,Indice],True);
  until TOBFC = nil;
end;

procedure TGestionConso.RecupCoefsPO (TOBLigne : TOB; VenteAchat : string; var CoefpaPr : double; var CoefPrPv : double);
begin
	RecupCoefs (TOBLIgne,VenteAchat,CoefPaPr,CoefPrPv);
end;

function NewConsoFromLigne (TOBLigne,TOBC : TOB) : boolean;
Var Part      : String;
    Part0     : String;
    Part1     : String;
    Part2     : String;
    Part3     : String;
    Part4     : String;
    Requete   : String;
    Annee     : word;
    Mois      : word;
    Jour      : word;
    Semaine   : integer;
    DateMouv  : TDateTime;
    Q         : TQuery;
    Nature    : String;
    NumMouv   : Double;
    TheRetour : TGncERROR;
    RatioVente : double;
    stprec,VenteAchat,TypeArticle : string;
    CoefPAPR,COEfPrPV : double;
Begin
	result := false;
  if TobLigne = nil then Exit;
  if TobLigne.GetValue('GL_AFFAIRE') = '' then exit;
  VenteAchat := GetInfoParPiece(TobLigne.GetValue('GL_NATUREPIECEG'), 'GPP_VENTEACHAT');

  Part  := TobLigne.GetValue('GL_AFFAIRE');
  Part0 := '';
  Part1 := '';
  Part2 := '';
  Part3 := '';
  Part4 := TobLigne.GetValue('GL_AVENANT');

  DateMouv := TobLigne.GetValue('GL_DATEPIECE');

  // Découpage du code Affaire
{$IFDEF BTP}
  BTPCodeAffaireDecoupe(Part,Part0,Part1,Part2,Part3, Part4, TaModif,false);
{$ELSE}
  CodeAffaireDecoupe(Part,Part0,Part1,Part2,Part3, Part4, TaModif,false);
{$ENDIF}

  // Découpage de la date
  DecodeDate(DateMouv, Annee, Mois, Jour);

  // Recherche du Numéro de semaine
  Semaine := NumSemaine(DateMouv);
  TypeArticle := TobLigne.GetValue('GL_TYPEARTICLE');
  // Recherche de la Nature du Mouvement si Prestation
  if TypeArticle = 'PRE' then
  Begin
    Requete := 'SELECT GA_ARTICLE,GA_LIBELLE,N1.BNP_TYPERESSOURCE,N1.BNP_LIBELLE FROM ARTICLE ART ' +
               'LEFT JOIN NATUREPREST N1 ON ART.GA_NATUREPRES=N1.BNP_NATUREPRES ' +
               'WHERE GA_TYPEARTICLE="PRE" AND GA_ARTICLE="'+TOBLigne.GetValue('GL_ARTICLE')+'"';
    Q := OpenSQL(Requete, True);
    if not Q.eof then
    begin
      Nature := Q.findfield('BNP_TYPERESSOURCE').asString ;
      ferme (Q);
    end else
    Begin
      ferme (Q);
      exit;
    end;
  end;

  if venteAchat = 'ACH' then
  begin
  	CoefPaPr := 0;
    CoefPrPv := 0;
  	RatioVente := GetRatio(TOBLigne, nil, trsVente);
    RecupCoefs (TOBLIgne,VenteAchat,CoefPaPr,CoefPrPv);
  end else
  begin
  	RatioVente := 1;
  	CoefPaPr := 0;
    CoefPrPv := 0;
    if TOBLigne.GetValue('GL_DPA') <> 0 then CoefPaPR := TOBLigne.GetValue('GL_DPR')/TOBLigne.GetValue('GL_DPA');
    if TOBLigne.GetValue('GL_DPR') <> 0 then CoefPRPV := TOBLigne.GetValue('GL_PUHT')/TOBLigne.GetValue('GL_DPR');
	end;

    // Calcul du Numéro de Mouvement
  TheRetour := GetNumUniqueConso (NumMouv);
  if TheRetour = gncAbort then
  BEGIN
    Exit;
  END;

  // Chargement de la TOB Conssomation

  TOBC.PutValue('BCO_AFFAIRE', TobLigne.GetValue('GL_AFFAIRE'));
  TOBC.PutValue('BCO_AFFAIRE0', Part0);
  TOBC.PutValue('BCO_AFFAIRE1', Part1);
  TOBC.PutValue('BCO_AFFAIRE2', Part2);
  TOBC.PutValue('BCO_AFFAIRE3', Part3);
  TOBC.PutValue('BCO_PHASETRA', '');
  TOBC.PutValue('BCO_MOIS', Mois);
  TOBC.PutValue('BCO_SEMAINE', Semaine);
  TOBC.PutValue('BCO_DATEMOUV', DateMouv);
//INT,ST,AUT,LOC
  if TypeArticle = 'PRE' then
  begin
    if Nature = 'SAL' then
      TOBC.PutValue('BCO_NATUREMOUV', 'MO')
    else if  Nature = 'ST' then
      TOBC.PutValue('BCO_NATUREMOUV', 'EXT')
    else if  Nature = 'INT' then
      TOBC.PutValue('BCO_NATUREMOUV', 'EXT')
    else if  Nature = 'AUT' then
      TOBC.PutValue('BCO_NATUREMOUV', 'EXT')
    else if  Nature = 'LOC' then
      TOBC.PutValue('BCO_NATUREMOUV', 'EXT')
    else if  Nature = 'MAT' then
      TOBC.PutValue('BCO_NATUREMOUV', 'MAT')
    else if  Nature = 'OUT' then
      TOBC.PutValue('BCO_NATUREMOUV', 'MAT');
  end else if (TypeArticle = 'MAR') or (TypeArticle = 'ARP') then
  Begin
    TOBC.PutValue('BCO_NATUREMOUV', 'FOU');
  end else if TypeArticle = 'FRA' then
  Begin
    TOBC.PutValue('BCO_NATUREMOUV', 'FRS');
  end;

  TOBC.PutValue('BCO_LIBELLE', TobLigne.GetValue('GL_LIBELLE'));
  TOBC.PutValue('BCO_TRANSFORME', '-');
  StPrec := TOBLigne.GetValue('GL_PIECEPRECEDENTE');

  TOBC.PutValue('BCO_NUMMOUV', NumMouv);
  TOBC.PutValue('BCO_INDICE',0);
  if TObLigne.getValue('GL_PRIXPOURQTE')= 0 then TOBLigne.putvalue('GL_PRIXPOURQTE',1);
  TOBC.PutValue('BCO_CODEARTICLE', TobLigne.GetValue('GL_CODEARTICLE'));
  TOBC.PutValue('BCO_ARTICLE', TobLigne.GetValue('GL_ARTICLE'));
  TOBC.PutValue('BCO_QUANTITE', TobLigne.GetValue('GL_QTERESTE') * RatioVente);
  TOBC.PutValue('BCO_QUALIFQTEMOUV', TobLigne.GetValue('GL_QUALIFQTEVTE'));
  TOBC.PutValue('BCO_FAMILLENIV1', TobLigne.GetValue('GL_FAMILLENIV1'));
  TOBC.PutValue('BCO_FAMILLENIV2', TobLigne.GetValue('GL_FAMILLENIV2'));
  TOBC.PutValue('BCO_FAMILLENIV3', TobLigne.GetValue('GL_FAMILLENIV3'));

  if (venteAchat = 'ACH') then
  begin
    TOBC.PutValue('BCO_DPA', TobLigne.GetValue('GL_PUHTNET')/(TOBLigne.getValue('GL_PRIXPOURQTE')));
    TOBC.PutValue('BCO_DPR', Arrondi(TOBC.GEtValue('BCO_DPA')*CoefPaPr,V_PGI.okdecP));
    TOBC.PutValue('BCO_PUHT',Arrondi(TOBC.GEtValue('BCO_DPR')*CoefPRPV,V_PGI.okdecP));
  end else
  begin
    TOBC.PutValue('BCO_DPA', (TobLigne.GetValue('GL_DPA')/(TOBLigne.getValue('GL_PRIXPOURQTE'))* RatioVente));
    TOBC.PutValue('BCO_DPR', (TobLigne.GetValue('GL_DPA')/(TOBLigne.getValue('GL_PRIXPOURQTE'))*CoefPaPr));
    TOBC.PutValue('BCO_PUHT', (TobLigne.GetValue('GL_DPR')/(TOBLigne.getValue('GL_PRIXPOURQTE'))*CoefPrPv));
  end;

  TOBC.PutValue('BCO_FACTURABLE', 'N');
  //
  if Part0 = 'W' then
  begin
    // Appel d'intervention
    TOBC.PutValue('BCO_FACTURABLE', 'A');
    TOBC.putValue('BCO_AFFAIRESAISIE',GetContrat(TobLigne.GetValue('GL_AFFAIRE')));
    TOBC.putValue('BCO_QTEFACTUREE',TOBC.GetValue('BCO_QUANTITE'));
  end;
  //
  TOBC.PutValue('BCO_VALIDE', '-');
  TOBC.PutValue('BCO_NATUREPIECEG', TobLigne.GetValue('GL_NATUREPIECEG'));
  TOBC.PutValue('BCO_SOUCHE', TobLigne.GetValue('GL_SOUCHE'));
  TOBC.PutValue('BCO_NUMERO', TobLigne.GetValue('GL_NUMERO'));
  TOBC.PutValue('BCO_INDICEG', TobLigne.GetValue('GL_INDICEG'));
  TOBC.PutValue('BCO_NUMORDRE', TobLigne.GetValue('GL_NUMORDRE'));
  CalculeLaLigne(TOBC);
  result := true;
end;

function EncodeRefPieceLivr (TOBL : TOB) : string;
begin
  Result := TOBL.GetValue('GL_NATUREPIECEG') + ';' + TOBL.GetValue('GL_SOUCHE') + ';'+
  					IntToStr(TOBL.GetValue('GL_NUMERO')) + ';' + IntToStr(TOBL.GetValue('GL_INDICEG')) + ';'+
  					IntToStr(TOBL.GetValue('GL_NUMORDRE')) + ';';
end;

function EncodeRefPieceLivr (refpieceStd : string) : string;
var Cledoc : R_CLEDOC;
begin
	Decoderefpiece(refPieceStd,Cledoc);
  Result := Cledoc.NaturePiece +';' + Cledoc.Souche + ';'+
  					IntToStr(Cledoc.NumeroPiece) + ';' + IntToStr(Cledoc.Indice) + ';'+
  					IntToStr(Cledoc.NumOrdre) + ';';
end;

procedure TGestionConso.CreeAssociation (TOBPhases,TOBL : TOB ; Nummouv : double = 0);
var UneTob : TOB;
begin
  if TOBL.GetValue('GL_AFFAIRE')='' then exit;
  if not TOBL.FieldExists('BLP_PHASETRA') then TOBL.AddChampSupValeur('BLP_PHASETRA','');
  UneTob := TOB.Create ('LIGNEPHASES',TOBPhases,-1);
  UneTOB.putValue('BLP_NATUREPIECEG',TOBL.GetValue('GL_NATUREPIECEG'));
  UneTOB.putValue('BLP_SOUCHE',TOBL.GetValue('GL_SOUCHE'));
  UneTOB.putValue('BLP_NUMERO',TOBL.GetValue('GL_NUMERO'));
  UneTOB.putValue('BLP_INDICEG',TOBL.GetValue('GL_INDICEG'));
  UneTOB.putValue('BLP_NUMORDRE',TOBL.GetValue('GL_NUMORDRE'));
  UneTOB.putValue('BLP_PHASETRA',TOBL.GetValue('BLP_PHASETRA'));
  UneTOB.putValue('BLP_NUMMOUV',NumMouv);
end;

procedure TGestionConso.GestionModifLivraison (TOBGenere,TOBCONSODEL : TOB);
var Indice,Ind : integer;
		TOBL : TOB;
    QQ : TQuery;
    TOBreceptions,TOBR,TOBLIVRAISONS,TOBLL,TOBREC,TOBPhases : TOB;
    SQl : string;
    PrevQteLiv,Nummouv : double;
    DateMouv : TdateTime;
    Annee,mois,jour : word;
    Semaine : integer;
    CoefPaPr,CoefPrPv : double;
begin
	TOBreceptions := TOB.Create('LES RECEPTIONS DE LA LIV',nil,-1);
	TOBLIVRAISONS := TOB.Create('LES LIVRAISONS A MODIF',nil,-1);
	TOBPhases := TOB.Create('LES Phases',nil,-1);
  TRY
  	if (TOBCONSODEL <> nil) and (TOBCONSODEL.detail.count >0) then
    begin
    	// Suppression da la ligne de conso livraison  + maj des lignes de conso liés aux receptions/facture fournisseurs
      for Indice := 0 to TOBConsodel.detail.count -1 do
      begin
      	prevQteLiv := 0;
      	nummouv := TOBCONSODEL.detail[Indice].getValue('CONSOSUP');
        Sql := 'SELECT BCO_QUANTITE FROM CONSOMMATIONS WHERE BCO_NUMMOUV='+
        			floattostr(nummouv);
        QQ := OpenSql (Sql,true,1,'',true);
        if not QQ.eof then
        begin
        	PrevQteLiv := QQ.findField('BCO_QUANTITE').asfloat;
        end;
        ferme (QQ);
        if prevQteLiv <> 0 then
        begin
          TOBR := TOB.Create ('LES RECEP / LIGNE LIV',TOBreceptions,-1);
          Sql := 'SELECT * FROM CONSOMMATIONS WHERE BCO_NUMMOUV IN (SELECT BCO_NUMMOUV FROM CONSOMMATIONS WHERE BCO_LIENVENTE='+
              		FloatToStr(nummouv)+') AND BCO_INDICE=0';
          TOBR.LoadDetailDBFromSQL ('CONSOMMATIONS',SQl,false);
          for ind := 0 to TOBR.detail.count -1 do
          begin
            TOBREC := TOBR.detail[Ind];
            TOBRec.putvalue('BCO_QTEVENTE',TOBRec.getvalue('BCO_QTEVENTE')-PrevQteliv);
            if TOBREC.GetValue('BCO_QTEVENTE')=0 then TOBREC.putValue('BCO_TRAITEVENTE','-');
          end;
          // les lignes de reception et/ou facture fournisseur liée à a la livraison
          ExecuteSql('DELETE from LIGNEPHASES WHERE BLP_NUMMOUV='+FloatToStr(nummouv));
          // Pas de test du retour car la ligne du document d'achat correspondant a pu déjà être supprimée
          ExecuteSql('DELETE FROM CONSOMMATIONS WHERE BCO_LIENVENTE='+FloatToStr(nummouv)+' AND BCO_INDICE >0');
          // la ligne de livraison sur chantier
          if ExecuteSql('DELETE from CONSOMMATIONS WHERE BCO_NUMMOUV='+
              FloatToStr(nummouv)) = 0 then V_PGI.IOError := oeUnknown;
        end;
      end;
    end;
    for Indice := 0 to TOBgenere.detail.count -1 do
    begin
      TOBL := TOBGenere.detail[Indice];
      if TOBL.getValue('GL_AFFAIRE') = '' then continue;
      if (TOBL.GetValue('GL_TYPELIGNE')='ART') then
      begin
      	if (TOBL.GetValue('BLP_NUMMOUV')<>0) and (TOBL.IsOneModifie (true)) then
        begin
        	prevQteLiv := 0;
          TOBLL := TOB.Create ('CONSOMMATIONS',TOBLIVRAISONS,-1);
          Sql := 'SELECT * FROM CONSOMMATIONS WHERE BCO_NUMMOUV='+
                 FloatToStr(TOBL.getvalue('BLP_NUMMOUV'));
          QQ := OpenSql (Sql,true,1,'',true);
          if not QQ.eof then
          begin
            TOBLL.selectDb ('',QQ);
            PrevQteLiv := TOBLL.getValue('BCO_QUANTITE');
            TOBLL.putValue('BCO_QUANTITE',TOBL.GetValue('GL_QTEFACT'));
            TOBLL.putValue('BCO_DATEMOUV',TOBL.GetValue('GL_DATEPIECE'));
            DateMouv := TOBL.GetValue('GL_DATEPIECE');
            // Mise à jour des éléments de date
            DecodeDate(DateMouv, Annee, Mois, Jour);
            // Recherche du Numéro de semaine
            Semaine := NumSemaine(DateMouv);
            //
            TOBLL.PutValue('BCO_MOIS', Mois);
            TOBLL.PutValue('BCO_SEMAINE', Semaine);

  	        // Recalcul PU et Montants
            CoefPaPr := 0;
            CoefPrPv := 0;
            if TOBL.GetValue('GL_DPA') <> 0 then CoefPaPR := TOBL.GetValue('GL_DPR')/TOBL.GetValue('GL_DPA');
            if TOBL.GetValue('GL_DPR') <> 0 then CoefPRPV := TOBL.GetValue('GL_PUHT')/TOBL.GetValue('GL_DPR');
            TOBLL.PutValue('BCO_DPA', (TOBL.GetValue('GL_DPA')/(TOBL.getValue('GL_PRIXPOURQTE'))));
            CalculeLaLigne(TOBLL,CoefPaPR,CoefPRPV);

            // Maj code affaire
            TOBLL.putValue('BCO_AFFAIRE',TOBL.GetValue('GL_AFFAIRE'));
            TOBLL.putValue('BCO_AFFAIRE0',Copy(TOBL.GetValue('GL_AFFAIRE'),1,1));
            TOBLL.putValue('BCO_AFFAIRE1',TOBL.GetValue('GL_AFFAIRE1'));
            TOBLL.putValue('BCO_AFFAIRE2',TOBL.GetValue('GL_AFFAIRE2'));
            TOBLL.putValue('BCO_AFFAIRE3',TOBL.GetValue('GL_AFFAIRE3'));
          end;
          ferme (QQ);
          if (not TOBLL.IsOneModifie (true)) then TOBLL.free;
          //
          TOBR := TOB.Create ('LES RECEP / LIGNE LIV',TOBreceptions,-1);
          TOBR.AddChampSupValeur('REFLIVR',TOBL.getValue('GL_NATUREPIECEG')+';'+inttostr(TOBL.getValue('GL_NUMERO')));
          //
          Sql := 'SELECT * FROM CONSOMMATIONS WHERE BCO_NUMMOUV IN '+
                 '(SELECT BCO_NUMMOUV FROM CONSOMMATIONS WHERE BCO_LIENVENTE='+
                 FloatToStr(TOBL.getvalue('BLP_NUMMOUV'))+')';
          //
          TOBR.LoadDetailDBFromSQL ('CONSOMMATIONS',SQL,false);
          if TOBR.detail.count > 0 then
          begin
            for ind := 0 to TOBR.detail.count -1 do
            begin
              TOBREC := TOBR.detail[ind];
              if (TOBREC.GetValue('BCO_LIENVENTE')=TOBL.getvalue('BLP_NUMMOUV')) then
              begin
                TOBREC.PutValue('BCO_QTEVENTE',TOBL.getValue('GL_QTEFACT'));
                TOBREC.PutValue('BCO_QUANTITE',TOBL.getValue('GL_QTEFACT'));
              end else if (TOBREC.GetValue('BCO_INDICE')=0) then
              begin
                TOBREC.PutValue('BCO_QTEVENTE',TOBREC.getValue('BCO_QTEVENTE')-PrevQteLiv+TOBL.getValue('GL_QTEFACT'));
                if TOBREC.GetValue('BCO_QTEVENTE')<>TOBREC.getValue('BCO_QUANTITE') then TOBREC.putValue('BCO_TRAITEVENTE','-');
              end;
            end;
            //
            ind := 0;
            repeat
              TOBRec := TOBR.detail[Ind];
              if (not TOBREC.IsOneModifie (true)) then TOBREC.free else inc(ind);
            until ind >= TOBR.detail.count;
          end;
          if TOBR.detail.count = 0 then TOBR.free;
          //
        end else
        begin
        	// forcement la on cree une ligne
      		if TOBL.GetValue('GL_QTEFACT') = 0 then continue;
          TOBLL := TOB.Create ('LA LIGNE LIV',TOBLIVRAISONS,-1);
          NumMouv := CreerConso(TOBL, TOBL.GetValue('BLP_PHASETRA'),TTcoCreat,TOBLL);
        	CreeAssociation (TOBPhases,TOBL,Nummouv);
      		TOBL.PutValue('BLP_NUMMOUV',Nummouv);
          TOBLL.setAllModifie(true);
          TOBPhases.setAllModifie(true);
          TOBLL.InsertDB (nil);
          TOBPhases.InsertOrUpdateDB;
          TOBL.UpdateDB(false);
        end;
      end;
    end;
    if (TOBReceptions.detail.count > 0) and (TOBReceptions.IsOneModifie (true)) then if not TOBreceptions.UpdateDB then V_PGI.ioerror := oeUnknown;
    if (TOBLIVRAISONS.detail.count > 0) and (TOBLIVRAISONS.IsOneModifie (true)) then if not TOBLIVRAISONS.UpdateDB then V_PGI.ioerror := oeUnknown;
    if (TOBPhases.detail.count > 0) then
    begin
    	TOBPhases.setallmodifie(true);
    	if not TOBPhases.UpdateDB then V_PGI.ioerror := oeUnknown;
    end;
    // gestion des lignes suppprimes
  FINALLY
  	TOBReceptions.free;
    TOBLIVRAISONS.free;
    TOBPhases.free;
  END;
end;

end.
