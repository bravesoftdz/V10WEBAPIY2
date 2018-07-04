unit BTPrepaLivr;

interface

Uses StdCtrls,
     Controls,
     Classes,
     Mul,
     M3FP,
     AglInit,
     Ent1,
{$IFDEF EAGLCLIENT}
     Maineagl,
{$ELSE}
     FE_Main,
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     utofAfBaseCodeAffaire,
     HTB97,
     UTOB,
     EntGC,
     FactComm,
     FactGrp,
{$IFDEF BTP}
     BTPUTIL,
{$ENDIF}
     UTOF,uEntCommun ;

const
	TexteMessage: array[1..2] of string 	= (
          {1}        'Confirmez-vous le traitement de préparation des besoins de chantier ?',
          {2}        'Rien à traiter pour la date de préparation indiquée !');

Type


  R_ParamTrait = record
    ArtDeb, ArtFin, DepDeb, DepFin : string;
    Fam1, Fam2, Fam3, Fam1_, Fam2_, Fam3_, DatDeb, DatFin : string;
  end;

  TGenPrepaLiv = class
    private
    ThePieces,TOBPieces : TOB;
    CompteRendu : boolean;
    procedure MiseAjourPieceprec(TOBpiece: TOB);
    procedure UpdateLignePrec(Cledoc: R_Cledoc);
    procedure BeforeGenerationPieces(TOBPieces: TOB);
    procedure TraitePiece(TOBP: TOB);
    function Controleparagraphe(TOBP : TOB;Debut,Niveau : integer; Var FinParagraphe : integer) : boolean;
    procedure DeleteParagraphe(TOBP: TOB; Debut, Niveau: integer);
    public
    constructor create;
    destructor destroy; override;
    procedure GenereThepieces ;
  end;

  TOF_BTGENPREP_MUL = Class (TOF_AFBASECODEAFFAIRE)
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
  end ;

  TOF_BTDATEPREP = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;


procedure GenereCommandeChantier;
procedure GenereCommandeChantierFromRea ( Parametres : R_ParamTrait);

implementation
uses FactUtil,FactVariante;

function DefiniRequeteLigFromParams (Parametres : R_ParamTrait) : string;
begin
  result := '';
  result := 'SELECT *,PIE.GP_REPRESENTANT,PIE.GP_REFINTERNE,PIE.GP_NUMADRESSEFACT,PIE.GP_NUMADRESSELIVR,PHA.BLP_PHASETRA FROM LIGNE';
  result := result +' LEFT JOIN PIECE PIE ON (PIE.GP_NATUREPIECEG=GL_NATUREPIECEG) and (PIE.GP_SOUCHE=GL_SOUCHE) and (PIE.GP_NUMERO=GL_NUMERO) and (PIE.GP_INDICEG=GL_INDICEG)';
  result := result +' LEFT JOIN LIGNEPHASES PHA ON (PHA.BLP_NATUREPIECEG=GL_NATUREPIECEG) and (PHA.BLP_SOUCHE=GL_SOUCHE) and(PHA.BLP_NUMERO=GL_NUMERO) and (PHA.BLP_INDICEG=GL_INDICEG) and (PHA.BLP_NUMORDRE=GL_NUMORDRE)';
  result := result + ' WHERE GL_NATUREPIECEG = "PBT" ' ;
  if (Parametres.ArtDeb <> '') and (Parametres.ArtFin <> '') then
      result := result + 'AND (GL_CODEARTICLE>="' + Parametres.ArtDeb + '" AND GL_CODEARTICLE<="' + Parametres.ArtFin + '") '
  else if Parametres.ArtDeb <> '' then
      result := result + 'AND GL_CODEARTICLE>="' + Parametres.ArtDeb + '" '
  else if Parametres.ArtFin <> '' then
      result := result + 'AND GL_CODEARTICLE<="' + Parametres.ArtFin + '" ';
  if Parametres.Fam1 <> '' then
      if Parametres.Fam1_ <> '' then
          result := result + 'AND (GL_FAMILLENIV1>="' + Parametres.Fam1 + '" AND GL_FAMILLENIV1<="' + Parametres.Fam1_ + '") '
          else
          result := result + 'AND GL_FAMILLENIV1>="' + Parametres.Fam1 + '" ';
  if Parametres.Fam2 <> '' then
      if Parametres.Fam2_ <> '' then
          result := result + 'AND (GL_FAMILLENIV2>="' + Parametres.Fam2 + '" AND GL_FAMILLENIV2<="' + Parametres.Fam2_ + '") '
          else
          result := result + 'AND GL_FAMILLENIV2="' + Parametres.Fam2 + '" ';
  if Parametres.Fam3 <> '' then
      if Parametres.Fam3_ <> '' then
          result := result + 'AND (GL_FAMILLENIV3>="' + Parametres.Fam3 + '" AND GL_FAMILLENIV3<="' + Parametres.Fam3_ + '") '
          else
          result := result + 'AND GL_FAMILLENIV3="' + Parametres.Fam3 + '" ';
  if Parametres.DepDeb  <> '' then
      if Parametres.DepFin  <> '' then
          result := result + 'AND (GL_DEPOT>="' + Parametres.DepDeb + '" AND GL_DEPOT<="' + Parametres.DepFin + '") '
          else
          result := result + 'AND GL_DEPOT="' + Parametres.DepDeb + '" ';
  result := result + 'AND GL_DATELIVRAISON>="' + USDateTime(strtodate(Parametres.DatDeb)) + '" ';
  result := result + 'AND GL_DATELIVRAISON<="' + USDateTime(StrToDate(Parametres.DatFin)) + '" ' ;
  result := result + 'AND GL_TYPEDIM<>"GEN" ' ;
//  result := result + 'AND GL_TYPELIGNE="ART" AND GL_QTERESTE > 0 AND GL_VIVANTE="X"';
(*   result := result + 'AND GL_TYPELIGNE="ART" '; // Plus besoin maintenant *)
// oooppps
//  result := result + 'AND ((GL_QTERESTE > 0 AND ((GL_TYPELIGNE="ART") OR (GL_TYPELIGNE="ARV"))) OR GL_TYPELIGNE <> "ART")';
//  result := result + 'AND ((GL_QTERESTE <> 0 AND ((GL_TYPELIGNE="ART") OR (GL_TYPELIGNE="ARV"))) OR GL_TYPELIGNE <> "ART")';
    //--- GUINIER ---
  result := result + 'AND (((GL_MTRESTE <> 0 OR GL_QTERESTE <> 0) ';

  Result := Result + 'AND ((GL_TYPELIGNE="ART") OR (GL_TYPELIGNE="ARV"))) OR ((GL_TYPELIGNE <> "ART") AND (GL_TYPELIGNE <> "ARV")) )';
  result := result + 'ORDER BY GL_NUMERO,GL_NUMLIGNE';
end;

procedure GenereCommandeChantier;
begin
  AGLLanceFiche('BTP','BTGENPREP_MUL','GP_NATUREPIECEG=PBT;GP_VIVANTE=X','','');
end;

procedure TOF_BTGENPREP_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTGENPREP_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTGENPREP_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTGENPREP_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTGENPREP_MUL.OnArgument (S : String ) ;
var CC : THValComboBox;
begin
  Inherited ;

  // gestion Etablissement (BTP)
  CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ;
  if CC<>Nil then PositionneEtabUser(CC) ;
  //
  //Gestion Restriction Domaine
  CC:=THValComboBox(GetControl('GP_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

end ;

procedure TOF_BTGENPREP_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTGENPREP_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTGENPREP_MUL.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_BTGENPREP_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
  Aff:=THEdit(GetControl('GP_AFFAIRE'));
  Aff0:=THEdit(GetControl('GP_AFFAIRE0'));
  Aff1:=THEdit(GetControl('GP_AFFAIRE1'));
  Aff2:=THEdit(GetControl('GP_AFFAIRE2'));
  Aff3:=THEdit(GetControl('GP_AFFAIRE3'));
  Aff4:=THEdit(GetControl('GP_AVENANT'));
  Tiers:=THEdit(GetControl('GP_TIERS'));
end;

{ BTDatePrep }


procedure TOF_BTDATEPREP.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTDATEPREP.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTDATEPREP.OnUpdate ;
begin
  Inherited ;
  if LaTob.getValue('DATELIMITE') < V_PGI.DateEntree then
  begin
    PGIBox (TraduireMemoire('La date de prise en compte est inférieure à la date du jour'),ecran.caption);
    TForm(Ecran).ModalResult:=0;
    exit;
  end;
  TheTob := laTob;
end ;

procedure TOF_BTDATEPREP.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTDATEPREP.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_BTDATEPREP.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTDATEPREP.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTDATEPREP.OnCancel () ;
begin
  Inherited ;
end ;

function GetDateFinPriseEnCompte (var DateFin : TdateTime) : boolean;
var TOBDate : TOB;
begin
  Result := true;
  DateFin := 0;
  TOBDate := TOB.Create ('The Date',nil,-1);
  TOBDate.AddChampSupValeur ('DATELIMITE',V_PGI.DateEntree,false);
  TRY
    TheTob := TOBDate;
    AGLLanceFiche('BTP','BTDATEPREP','','','') ;
    if TheTob = nil then BEGIN Result := false; Exit; END;
    DateFin := TheTob.getValue('DATELIMITE');
  FINALLY
    TOBDate.free;
    TheTob := nil;
  END;
end;

procedure AGLGenereCommandesChantier(parms : array of variant; nb : integer);
var Numero,Indice,IndiceP,NBArt : integer;
    Naturepiece,Souche,DatePiece : string;
    TheForm : Tform;
    Datefin : TDateTime;
    TOBL: TOB;
    Requete : string;
    QQ : Tquery;
    PrepaLivraison : TGenPrepaLiv;
begin
  TheForm:=TForm(Longint(Parms[0]));
  PrepaLivraison := TGenPrepaLiv.create ;
  PrepaLivraison.CompteRendu := true;
  TRY
    if (PGIAsk (traduireMemoire(TexteMessage[1]), TheForm.Caption)=mrYes) then
    begin
      if GetDateFinPriseEnCompte (DateFin) then
      begin
        SourisSablier;
        for indice:=0 to TFMul(TheForm).fliste.nbSelected-1 do
        begin
          TFMul(TheForm).fliste.GotoLeBookmark(indice);
          NaturePiece:=TFMul(TheForm).fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
          DatePiece:=TFMul(TheForm).fliste.datasource.dataset.FindField('GP_DATEPIECE').AsString;
          Souche:=TFMul(TheForm).fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
          Numero:=TFMul(TheForm).fliste.datasource.dataset.FindField('GP_NUMERO').AsInteger;
          IndiceP:=TFMul(TheForm).fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger;

          // Nouvelle requete corrige par LS
          Requete := 'SELECT LIG.*,PIE.GP_REPRESENTANT,PIE.GP_REFINTERNE,PIE.GP_NUMADRESSEFACT,PIE.GP_NUMADRESSELIVR,PHA.BLP_PHASETRA FROM LIGNE LIG '+
                     'LEFT JOIN PIECE PIE ON (PIE.GP_NATUREPIECEG=GL_NATUREPIECEG) and (PIE.GP_SOUCHE=GL_SOUCHE) '+
                     'and(PIE.GP_NUMERO=GL_NUMERO) and (PIE.GP_INDICEG=GL_INDICEG) '+
                     'LEFT JOIN LIGNEPHASES PHA ON (PHA.BLP_NATUREPIECEG=GL_NATUREPIECEG) and (PHA.BLP_SOUCHE=GL_SOUCHE) and '+
                     '(PHA.BLP_NUMERO=GL_NUMERO) and (PHA.BLP_INDICEG=GL_INDICEG) and (PHA.BLP_NUMORDRE=GL_NUMORDRE) WHERE '+
                     'GL_NATUREPIECEG="'+NaturePiece+'" AND GL_SOUCHE="'+Souche+'" AND GL_NUMERO='+IntToStr(Numero)+' AND GL_INDICEG='+ IntToStr(IndiceP)+'AND '+
                     '((GL_TYPELIGNE IN ("ART","ARV")  AND (GL_QTERESTE<>0 OR GL_MTRESTE<>0) AND (GL_DATELIVRAISON <= "'+UsDateTime(DateFin)+'")) OR ( NOT GL_TYPELIGNE IN ("ART","ARV"))) ORDER BY GL_NUMLIGNE';

          QQ := OpenSql (Requete,true,-1,'',true);
          if not QQ.eof then
          begin
            PrepaLivraison.TOBPieces.LoadDetailDB ('LIGNE','','',QQ,true,true);
          end;
          ferme (QQ);
        end;
        NbArt := 0;
        if PrepaLivraison.TOBPieces.detail.count > 0 then
        begin
          for Indice := 0 to PrepaLivraison.TOBPieces.detail.count -1 do
          begin
            TOBL := PrepaLivraison.TOBPieces.detail[Indice];
            // Gestion Des frais detaillé dans Prévisions
            if isVariante (TOBL) then SetTypeLigne(TOBL,false);
            if TOBL.GetValue('GL_TYPELIGNE')='ART' Then Inc(NbArt);
            TOBL.putvalue('GL_PIECEPRECEDENTE',EncodeRefPiece(TOBL));
            TOBL.putvalue('GL_PIECEORIGINE',EncodeRefPiece(TOBL));
            TOBL.addChampSupValeur('ORIGINE',EncodeRefPiece(TOBL,0,false));
          end;
          if NbArt > 0 then Transactions(PrepaLivraison.GenereThepieces,1)
                       else PGIBox (traduireMemoire(TexteMessage[2]), TheForm.Caption);
        end else
        begin
          PGIBox (traduireMemoire(TexteMessage[2]), TheForm.Caption);
        end;
      end;
    end;
  FINALLY
  PrepaLivraison.Free;;
  SourisNormale ;
  END;
end;

procedure GenereCommandeChantierFromRea ( Parametres : R_ParamTrait);
var Requete : string;
    QQ : Tquery;
    PrepaLivraison : TGenPrepaLiv;
    indice,NBArt : integer;
    TOBL : TOB;
begin
  NBArt := 0;
  PrepaLivraison := TGenPrepaLiv.create ;
  PrepaLivraison.CompteRendu := false;
  TRY
    Requete := DefiniRequeteLigFromParams (Parametres);
    QQ := OpenSql (Requete,true,-1,'',true);
    if not QQ.eof then
    begin
      PrepaLivraison.TOBPieces.LoadDetailDB ('LIGNE','','',QQ,true,true);
    end;
    ferme (QQ);
    if PrepaLivraison.TOBPieces.detail.count > 0 then
    begin
      for Indice := 0 to PrepaLivraison.TOBPieces.detail.count -1 do
      begin
        TOBL := PrepaLivraison.TOBPieces.detail[Indice];
        // Gestion Des frais detaillé dans Prévisions
        if isVariante (TOBL) then SetTypeLigne(TOBL,false);
        // --
        if TOBL.GetValue('GL_TYPELIGNE')='ART' Then Inc(NbArt);
        TOBL.putvalue('GL_PIECEPRECEDENTE',EncodeRefPiece(TOBL));
        TOBL.putvalue('GL_PIECEORIGINE',EncodeRefPiece(TOBL));
        TOBL.addChampSupValeur('ORIGINE',EncodeRefPiece(TOBL,0,false));
      end;
      if NbARt > 0 then Transactions(PrepaLivraison.GenereThepieces,1);
    end;
  FINALLY
    PrepaLivraison.free;
  END;
end;

{ TGenPrepaLiv }

constructor TGenPrepaLiv.create;
begin
  TOBPieces := TOB.Create ('LES PIECES',nil,-1);
  ThePieces := TOB.Create ('LES PIECES GENEREES',nil,-1);
end;

destructor TGenPrepaLiv.destroy;
begin
  TOBPieces.free;
  ThePieces.free;
  inherited;
end;

function TGenPrepaLiv.Controleparagraphe(TOBP : TOB;Debut,Niveau : integer; Var FinParagraphe : integer) : boolean;
var Indice,NivSuivant,FinSUiv : integer;
		TOBL : TOB;
begin
	result := false;
  Indice := Debut;
  repeat
  	TOBL := TOBP.detail[Indice];
  	if IsFinParagraphe (TOBL,Niveau) then
    BEGIN
    	FinParagraphe := Indice;
    	break;
    END;
    if IsDebutParagraphe (TOBL,Niveau) then BEGIN Inc(Indice); continue; END;
    if IsDebutParagraphe (TOBL) then
    begin
      NivSuivant := TOBL.GetValue('GL_NIVEAUIMBRIC');
      if Controleparagraphe ( TOBP,Indice,NivSuivant,FinSuiv) then
      begin
        // il y a bien qq chose a traiter
      	result := true;
        Indice := FinSuiv;
      end else
      begin
        Indice := FinSuiv;
      end;
    end;
    if TOBL.getValue('GL_NIVEAUIMBRIC') > Niveau then BEGIN Inc(Indice); continue; END;
    if TOBL.GetValue('GL_TYPELIGNE') <> 'ART' Then BEGIN Inc(Indice); Continue; END; // on saute les lignes de commentaires
    if TOBL.GetValue('GL_TYPEARTICLE') <> 'PRE' then
    begin
    	// s'il y a une ligne d'un autre type que prestation ..on est sur qu'il n'est pas vide
    	result := true;
      inc(Indice);
    end else
    if (TOBL.GetValue('GL_TYPEARTICLE') = 'PRE') and (RenvoieTypeRes(TOBL.GetValue('GL_ARTICLE'))<>'SAL') then
    begin
    	// C'est une prestation de type autre que salarié ..donc a traiter
    	result := true;
      inc(Indice);
    end else Inc(Indice);
  until Indice >= TOBP.detail.count;
end;

procedure TGenPrepaLiv.DeleteParagraphe (TOBP : TOB;Debut,Niveau : integer);
var Indice : integer;
		TOBl : TOB;
    StopIt : boolean;
begin
	Indice := Debut;
  StopIt := false;
	Repeat
  	TOBL := TOBP.detail[Indice];
    if IsFinParagraphe (TOBL,Niveau) then
    begin
    	StopIt := true;
    end;
    TOBL.free;
  until (Indice >= TOBP.detail.count) or (StopIt);
end;

procedure TGenPrepaLiv.TraitePiece(TOBP : TOB);
var indice,FinParag : integer;
		TOBL : TOB;
begin
	Indice := 0;
  repeat
    TOBL := TOBP.detail[Indice];
    if TOBL.getValue('GL_NIVEAUIMBRIC') = 1 then
    begin
      if IsDebutParagraphe (TOBL,1) then
      begin
        if not Controleparagraphe(TOBP,Indice,1,FinParag) then
        begin
          DeleteParagraphe (TOBP,Indice,1);
        end else
        begin
        	Indice := FinParag; // saute a la fin du paragraphe traité
        end;
      end else
      begin
        inc(Indice);
      end;
    end else inc(Indice);
  until Indice >= TOBP.detail.count;
end;

procedure TGenPrepaLiv.BeforeGenerationPieces (TOBPieces : TOB);
begin
	TraitePiece(TOBPieces);
end;

procedure TGenPrepaLiv.GenereThepieces;
var Indice : integer;
    TOBpiece : TOB;
begin
  BeforeGenerationPieces (TOBPieces);
  if CreerPiecesFromLignes(TobPieces,'CHANTOCBT',Idate1900,CompteRendu,false,ThePieces) then
  begin
    for indice := 0 to ThePieces.detail.count -1 do
    begin
      TOBPiece := ThePieces.detail[Indice];
      MiseAjourPieceprec (TOBpiece);
    end;
  end;
end;

procedure TGenPrepaLiv.UpdateLignePrec (Cledoc : R_Cledoc);
var chaine : string;
begin
  chaine := 'UPDATE LIGNE SET GL_QTERESTE=0, GL_MTRESTE=0 WHERE GL_NATUREPIECEG="'+cledoc.NaturePiece+'" AND ' +
            'GL_SOUCHE="'+ Cledoc.Souche+'" AND ' +
            'GL_NUMERO='+inttostr(cledoc.NumeroPiece)+' AND ' +
            'GL_NUMORDRE='+inttostr(cledoc.numordre);
  ExecuteSql (chaine);
end;

procedure TGenPrepaLiv.MiseAjourPieceprec (TOBpiece : TOB);
var Indice : integer;
    TobPrecTraite : TOB;
    TOBL : TOB;
    cledoc : R_CleDoc;
    TOBD : TOB;
begin
  TOBPrecTraite := TOB.create ('LES DOCUMENTS TRAITES',nil,-1);
  TRY
  	//gcvoirtob(tobpiece);
    // Mise à jour des lignes
    for Indice := 0 to TOBpiece.detail.count -1 do
    begin
      TOBL := TOBpiece.detail[Indice];
      if TOBL.GetValue('GL_PIECEPRECEDENTE') <> '' then
      begin
         DecodeRefPiece (TOBL.GetValue('GL_PIECEPRECEDENTE'),Cledoc);
         TOBD := TOBPrecTraite.findFirst (['GP_NATUREPIECEG','GP_SOUCHE','GP_NUMERO','GP_INDICEG'],
                                     [cledoc.NaturePiece,cledoc.Souche,cledoc.NumeroPiece,cledoc.indice],true);
         if TOBD = nil then
         begin
           TOBD := TOB.create ('PIECE',TobPrecTraite,-1);
           TOBD.putvalue ('GP_NATUREPIECEG',cledoc.Naturepiece);
           TOBD.putvalue ('GP_SOUCHE',cledoc.Souche);
           TOBD.putvalue ('GP_NUMERO',cledoc.NumeroPiece);
           TOBD.putvalue ('GP_INDICEG',cledoc.Indice);
           TOBD.LoadDB (true);
         end;
         UPdateLignePrec (Cledoc);
      end;
    end;
    // mise à jour des pièces
    for Indice := 0 to TOBPrecTraite.detail.count -1 do
    begin
      TOBD := TOBPrecTraite.detail[Indice];
      if not ISAliveLine (TOBD) then
      begin
        TOBD.PutValue('GP_VIVANTE','-');
        TOBD.UpdateDB;
      end;
    end;
  FINALLY
    TOBPrecTraite.free;
  END;
end;

Initialization
  RegisterAglProc( 'GenereCommandesChantier',True,0,AGLGenereCommandesChantier);
  registerclasses ( [ TOF_BTGENPREP_MUL, TOF_BTDATEPREP ] ) ;

end.


