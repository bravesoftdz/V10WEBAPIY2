{***********UNITE*************************************************
Auteur  ...... : Dominique BROSSET
Créé le ...... : 08/02/2001
Modifié le ... : 10/07/2001
Description .. : Edition différée des documents
Mots clefs ... : EDITION;DIFFEREE;DOCUMENT
*****************************************************************}
unit UTofBTEditDocDiff;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, HDimension,UTOM,AGLInit,
      Utob,HDB,Messages,HStatus,uTofAfBaseCodeAffaire,
{$IFDEF EAGLCLIENT}
      UtileAGL,eMul,MaineAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fiche, mul, db,DBGrids,   Fe_Main,EdtEtat, EdtDoc,
      uPDFBatch,
{$IFNDEF V530}
      EdtREtat, EdtRDoc,
{$ENDIF}      
{$ENDIF}
{$IFDEF BTP}
      BTPUtil, Paramsoc,UtilReglementAffaire,
{$ENDIF}
{$IFDEF NOMADE}
      UtilPOP,
{$ENDIF}
			hPDFPrev,hPDFViewer,
      {$IFNDEF V530} uRecupSQLModele, {$ENDIF}
      M3VM, M3FP, Hqry, EntGC, FactComm, FactUtil,
      FactTob,Ent1,uEntCommun,UtilTOBPiece,
      Dialogs,      // TPrintDialog
      UDemandePrix,
      Windows,
      printers,
      BTFactImprTOB
 ;

Type
     TOF_BTEditDocDiff_Mul = Class (TOF_AFBASECODEAFFAIRE)
        procedure OnArgument (stArgument : string); override;
        procedure OnLoad; Override;
        procedure LanceEdition ;
        procedure EditeLesDocs (TobPiece : TOB);
        procedure NomsChampsAffaire( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ; override ;
//        Function  NomCorrect ( NomPDF : String ) : Boolean ;
  private

    //IsAvanc : boolean;
    VenteAchat : String;
    Statut		 : String;
    //
    NaturePieceg : THValcomboBox;

    procedure ControleChamp(Champ, Valeur: String);
    procedure ModificationModeleBTP(var OkEtat, BApercu: boolean; var Modele: string; var NbCopies: integer; EditDirect : boolean; LibelleAff : string);
    procedure SetInfoPieces(TobP: TOB; ETATSTANDARD: string; OKETAT: boolean; NbExemplaire: integer);
    procedure RenseignePiece(TobP: TOB; var BApercu: boolean; var Choix : string; var Modele: string; EditDirect: boolean; var Nbcopies: integer);
    function GetInfoRecapFromModele(TobPiece: TOB): boolean;
    procedure EditionDirecte(TOBP : TOB; StModele : string; ImprViaTOB : TImprPieceViaTOB; BDuplicata, Avancement : boolean);
    procedure EditionNonDirecte(TOBP : TOB; Cledoc : R_CleDoc;StModele, StSQl: string; BDuplicata: boolean);
    end ;

procedure AGLEditDocDiff_mul(parms:array of variant; nb: integer ) ;

const
// libellés des messages
TexteMessage: array[1..3] of string 	= (
          {1}  'Aucun élément sélectionné'
          {2} ,'Changement d''imprimé.' + #13 + 'Les documents suivants s''éditeront d''après le modèle '
          {2} ,'Trop de documents sélectionnés. Veuiller réduire la sélection'
              );

implementation
//
uses  FactureBtp,
      TntClasses,
      TntWideStrings, Spin;

procedure TOF_BTEditDocDiff_Mul.OnArgument (stArgument : string) ;
var CC      : THValComboBox;
    x				: integer;
    Critere : String;
    ValMul	: String;
    ChampMul: String;
begin
fMulDeTraitement  := true;
inherited ;

  fTableName := 'PIECE';

  if assigned(Getcontrol('GP_NATUREPIECEG')) then  NaturePieceG := THValComboBox(Ecran.FindComponent('GP_NATUREPIECEG'));

  Repeat
    Critere:=uppercase(ReadTokenSt(stArgument)) ;
    if Critere<>'' then
        begin
        x:=pos('=',Critere);
        if x<>0 then
           begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));
           end
        else
           ChampMul := Critere;
        ControleChamp(ChampMul, ValMul);
        end;
  until  Critere='';

  if statut <> '' then setcontrolText('AFF_STATUTAFFAIRE', Statut);

  if VenteAchat <> '' then
  begin
  if VenteAchat = 'STOCK' then
    NaturePieceG.Plus := 'AND (GPP_NATUREPIECEG="TEM" OR GPP_NATUREPIECEG="TRE" OR GPP_NATUREPIECEG="EEX" OR GPP_NATUREPIECEG="SEX")'
  else
    NaturePieceG.Plus := 'AND GPP_VENTEACHAT="' + VenteAchat + '"';
  end;

{$IFDEF NOMADE}
  if VenteAchat = 'VEN' then
  begin
    NaturePieceG.Plus := NaturePieceG.Plus + GetNaturePOP('GPP_NATUREPIECEG');
  end;
{$ENDIF} // NOMADE

  //*** Blocages des natures de pièces autorisées en affaires
  if (ctxAffaire in V_PGI.PGIContexte) and (VenteAchat='ACH') then
  begin
    if ctxScot in V_PGI.PGIContexte then
    begin
      NaturePieceG.Value := 'FF';
      NaturePieceG.Enabled := False;
    end
  else
  begin
    NaturePieceG.Plus := NaturePieceG.Plus + ' AND ((GPP_NATUREPIECEG="CF") or (GPP_NATUREPIECEG="BLF") or (GPP_NATUREPIECEG="FF"))';
    SetConTroltext('XX_WHERE',' AND GP_NATUREPIECEG IN ("CF","BLF","FF")');
  end;
  end;
  // **** Fin affaire ***

  //if not(ctxScot in V_PGI.PGIContexte) then   setControlproperty ('GP_NATUREPIECEG','Plus',stplus);

  if VenteAchat = 'VEN' then
  begin
    NaturePieceG.Value := 'CC';
    SetControlCaption ('TGP_TIERS', 'Client de');
  end
  else
  begin
    setControlproperty ('GP_TIERS','DataType','GCTIERSFOURN');
    setControlproperty ('GP_TIERS_','DataType','GCTIERSFOURN');
    SetControlText ('TGP_TIERS', 'Fourn. de');
    NaturePieceG.Value := 'CF';
  end;

  CC:=THValComboBox(GetControl('GP_DOMAINE'));
  if CC<>Nil then PositionneDomaineUser(CC) ;

  CC:=THValComboBox(GetControl('GP_ETABLISSEMENT'));
  if CC<>Nil then PositionneEtabUser(CC);

{$IFDEF BTP}
//SetControlVisible('PCOMPLEMENT',False);
  if VenteAchat = 'VEN' then
    begin
    NaturePieceG.Visible := False;
    end;
{$ENDIF}

  if Not(ctxAffaire in V_PGI.PGIContexte) and Not(ctxGCAFF in V_PGI.PGIContexte) then
   begin
   SetControlVisible ('TGP_AFFAIRE1', False);
   SetControlVisible ('GP_AFFAIRE1', False);
   SetControlVisible ('GP_AFFAIRE2', False);
   SetControlVisible ('GP_AFFAIRE3', False);
   SetControlVisible ('GP_AVENANT', False);
   end;

  if not VH_GC.GCMultiDepots then SetControlCaption('TGP_DEPOTDEST','Etablis. destinataire');

//uniquement en line
//SetControlVisible ('TGP_ETABLISSEMENT', False);
//SetControlVisible ('GP_ETABLISSEMENT', False);
//
  //FV1 - 14/03/2017 : FS#2431 - Viviane - Achats / Edition différée décocher "Edition mode direct"
  if VenteAchat = 'VEN' then
  begin
    SetControlChecked('EDITDIRECT',True);
    SetControlVisible('EDITDIRECT',True);
  end
  else
  begin
    SetControlChecked('EDITDIRECT',False);
    SetControlVisible('EDITDIRECT',False);
  end;

end;

Procedure TOF_BTEditDocDiff_Mul.ControleChamp(Champ : String;Valeur : String);
Begin
	if Champ='STOCK' then VenteAchat := Champ;
	if Champ='VEN' then VenteAchat := Champ;
  if Champ='ACH' then VenteAchat := Champ;

  if champ='COTRAITANCE' then Setcontroltext('XX_WHERE', 'AND AFF_MANDATAIRE IN ("X", "-")')
  else if champ ='SOUSTRAITANCE' then SetControltext('XX_WHERE', 'AND SSTRAITE > 0');

  //Modif FV pour gestion des type d'affaire dans les mul de recherches
  //if Champ='STATUT' then statut := Valeur;
  if Champ='STATUT' then
  Begin
    if Valeur = 'APP' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'W')
      else if assigned(GetControl('AFFAIRE0')) then
      begin
        SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("W")');
        SetControlText('AFFAIRE0', 'W');
      end;
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Appel');
      SetControlText('TGP_AFFAIRE', 'Code Appel');
    end
    else if Valeur = 'INT' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'I')
      else if assigned(GetControl('AFFAIRE0')) then
      begin
        SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("I")');
        SetControlText('AFFAIRE0', 'I');
      end;
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Contrat');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères contrat');
      SetControlText('TGP_AFFAIRE', 'Code Contrat');
    End
    else if Valeur = 'AFF' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'A')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'A');
      SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("A", "")');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Chantier');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Chantier');
      SetControlText('TGP_AFFAIRE', 'Code Chantier');
    end
    else if Valeur = 'GRP' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then
        SetControlText('XX_WHERE', ' AND AFF_AFFAIRE0 IN ("A","W")')
      else if assigned(GetControl('AFFAIRE0')) then
        SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("A","W")');
    end
    else if Valeur = 'PRO' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'P')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'P');
      SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("P")');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel d''offre');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Appel d''offre');
      SetControlText('TGP_AFFAIRE', 'Code Appel d''offre');
    end
    else
    Begin
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Affaire');
      SetControlText('TGP_AFFAIRE', 'Code Affaire');
    end;
  end;

	if CHAMP='INDIRECTE' then
  begin
    SetControlChecked('EDITDIRECT',false);
    SetControlVisible('EDITDIRECT',false);
  end else
	if CHAMP='VIVANTE' then
  begin
    SetControlTEXT('GP_VIVANTE','X');
  end;
end;

Procedure TOF_BTEditDocDiff_Mul.NomsChampsAffaire ( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ;
BEGIN
Aff:=THEdit(GetControl('GP_AFFAIRE'));
Aff0:=THEdit(GetControl('AFFAIRE0')) ;
Aff1:=THEdit(GetControl('GP_AFFAIRE1')) ;
Aff2:=THEdit(GetControl('GP_AFFAIRE2')) ;
Aff3:=THEdit(GetControl('GP_AFFAIRE3')) ;
Aff4:=THEdit(GetControl('GP_AVENANT'))  ;
END ;

procedure TOF_BTEditDocDiff_Mul.LanceEdition ;
var F : TFMul ;
    TobPiece : TOB ;
    iInd,II : integer ;
    StSql       : string;
    StSqlTrunc  : String;
    stWhere     : string;
    stNature    : string ;
    TSql : TQuery ;
    Indice : integer;
    cledoc : R_CLEDOC;
begin

  F:=TFMul(Ecran);
  if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
  begin
  {$IFNDEF EAGLCLIENT}
    if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
  {$ENDIF}
    HShowMessage('0;'+F.Caption+';'+TexteMessage[1]+';W;O;O;O;','','') ;
    exit;
  end;

  TobPiece := TOB.Create ('', Nil, -1) ;
  if TobPiece = nil then exit ;

  if F.FListe.AllSelected then
  begin
//    stWhere := RecupWhereCritere(F.Pages);
    StSql := THSQLMemo(GetControl('Z_SQL')).Lines.Text;
    StSql := Stringreplace(Trim((StSql)),'''','"',[rfReplaceAll]);
    //II := Pos('FROM',StSql);
    //StSqlTrunc := Copy(StSql,II,length(StSql));
    //II := Pos('ORDER BY',StSqlTrunc);
    //StSqlTrunc := Copy(StSqlTrunc,1,II-1);
    //StSql := 'SELECT count(GP_NUMERO) as nbr ' + stSqlTrunc;
    TSql := OpenSQL (StSql, True,-1,'',True);
    //if (not TSql.Eof) and (TSql.FindField ('nbr').AsInteger < 2000) then
    //begin
    //  Ferme (TSql);
    //  TSql := OpenSQL ('SELECT * ' + stSqlTrunc, True);
    //  if not TSql.EOF then TobPiece.LoadDetailDB('PIECE', '', '', TSql, False) else TobPiece := nil;
    //  Ferme (TSql) ;
    //end
    //else
    if not TSql.EOF then
      TobPiece.LoadDetailDB('PIECE', '', '', TSql, False)
    else
    begin
      TobPiece := nil;
      HShowMessage('0;'+TFMUL(Ecran).Caption+';'+TexteMessage[3]+';W;O;O;O;', '', '');
    end;
    Ferme (TSql);
  end
  else
  begin
    stWhere := '';
    stNature := NaturePieceG.value;
    for iInd:=0 to F.FListe.NbSelected-1 do
    begin
      F.FListe.GotoLeBOOKMARK(iInd);
{$IFDEF EAGLCLIENT}
      F.Q.TQ.Seek(F.FListe.Row-1) ;
{$ENDIF}
      if stnature <> '' then cledoc.NaturePiece := stNature
			                  //else cledoc.NaturePiece := F.Q.FindField('GP_SOUCHE').AsString;
      //FV1 - 19/01/2017 : FS#2161 - ECAL - Dans les achats et Editions différées, l'impression ne marche pas
                        else cledoc.NaturePiece := F.Q.FindField('GP_NATUREPIECEG').AsString;
      cledoc.Souche := F.Q.FindField('GP_SOUCHE').AsString;
      cledoc.NumeroPiece := F.Q.FindField('GP_NUMERO').asInteger;
      cledoc.Indice := F.Q.FindField('GP_INDICEG').asInteger;
      if not AlertDemandePrix (cledoc,TttPAcceptation) then continue;
      //
      if stWhere <> '' then stWhere := stWhere + ' OR ';

      stWhere := stWhere + '(GP_SOUCHE="'  + F.Q.FindField('GP_SOUCHE').AsString + '"' +
                       ' AND GP_NUMERO=' + F.Q.FindField('GP_NUMERO').AsString +
                       ' AND GP_INDICEG=' + F.Q.FindField('GP_INDICEG').AsString + ')';
    end;
    if StWhere <> '' then
    begin
       TSql := OpenSql ('SELECT * FROM PIECE WHERE GP_NATUREPIECEG="' + cledoc.NaturePiece + '" AND (' + stWhere + ')', True);
       if not TSql.EOF then TobPiece.LoadDetailDB('PIECE', '', '', TSql, False);
       Ferme (TSql);
    end;
  end;

  if TobPiece.Detail.Count > 0 then
  begin
    EditeLesDocs (TobPiece) ;
    if GetControlText ('GP_EDITEE') = '-' then
    begin
    	for Indice := 0 to TOBpiece.detail.count -1 do
      begin
				TobPiece.detail[Indice].UpdateDB(false);
      end;
    end;
  end;

  if F.FListe.AllSelected then F.FListe.AllSelected:=False else F.FListe.ClearSelected;
  F.bSelectAll.Down := False ;
  TobPiece.free ;
end;

(*
Function TOF_BTEditDocDiff_Mul.NomCorrect ( NomPDF : String ) : Boolean ;
Var FFText : TextFile ;
BEGIN
Result:=False ;
if NomPDF='' then Exit ;
if Pos('.PDF',NomPDF)<=0 then Exit ;
AssignFile(FFText,NomPDF) ; {$i-} Rewrite(FFText) ; {$i+}
if IoResult=0 then
   BEGIN
   CloseFile(FFText) ; DeleteFile(NomPDF) ;
   Result:=True ;
   END ;
END ;
*)

function TOF_BTEditDocDiff_Mul.GetInfoRecapFromModele(TobPiece : TOB) : boolean;
var SQl : string;
		Q : TQuery;
begin
  result := false;
  SQL:='SELECT BPD_IMPTABTOTSIT FROM BTPARDOC WHERE BPD_NATUREPIECE="'+TOBPiece.getValue('GP_NATUREPIECEG')+
  		 '" AND BPD_SOUCHE="'+TOBPiece.getValue('GP_SOUCHE')+
  		 '" AND BPD_NUMPIECE='+IntToStr(TOBPiece.getValue('GP_NUMERO')) ;
  Q:=OpenSQL(SQL,TRUE) ;
  if Not Q.EOF then
  begin
  	result:=(Q.Fields[0].AsString='X');
  end else
  Begin
    Ferme(Q);
    SQL:='SELECT BPD_IMPTABTOTSIT FROM BTPARDOC WHERE BPD_NATUREPIECE="'+TOBPiece.getValue('GP_NATUREPIECEG')+
    		 '" AND BPD_SOUCHE="'+TOBPiece.getValue('GP_SOUCHE')+'" AND BPD_NUMPIECE=0' ;
    Q:=OpenSQL(SQL,TRUE) ;
    if Not Q.EOF then result:=(Q.Fields[0].AsString='X')
  End;
  Ferme(Q) ;
end;

procedure TOF_BTEditDocDiff_Mul.EditeLesDocs (TobPiece : TOB) ;
var iInd, iNbExemplaire,iNbPiece : integer;
    CleDoc : R_CleDoc;
    //stModeleNonDirect
    stSql                 : string;
    stCle                 : string;
    stModele              : string;
    //NomPDF                : string;
    SQLOrder              : string;
    TL : TList;
    TT : TStrings;
    BApercu               : Boolean;
    BAp                   : Boolean;
    BDuplicata            : Boolean;
    //BFichierPDF         : Boolean;
    OldQRThread           : Boolean;
    OldQR                 : Boolean;
    BEditDirect           : boolean;
    //CFichierPDF         : TCheckBox;
    //ENomPDF             : THEdit;
    Pages : TpageControl;
    ImprViaTOB : TImprPieceViaTOB;
    Nbexemplaires : integer;
    recapSituations : boolean;
    TheModele             : string;        
    ModeleR               : string;
    TOBP                  : TOB;
    UneTOB : TOB;
    BEditionDirect : boolean;
    ChoixDirect           : string;
    ChoixNonDirect        : string;
    Choix                 : string;
    ModeleDirect          : string;
    ModeleNonDirect       : string;
  	ChoixAvancNonDirect   : string;
  	ModeleAvancNonDirect : string;
  	ChoixAvancDirect      : string;
  	ModeleAvancDirect     : string;
    OldprintDialog        : Boolean;
    Avancement            : boolean;
    PrintD                : TPrintDialog;
    aborted : boolean;
    OKEtatAvancNonDirect  : Boolean;
    OKEtatAvancDirect     : boolean;
    StMsg                 : String;
begin
  V_PGI.ioerror := OeOk;
  BEditDirect := TCheckBox(GetControl ('EDITDIRECT')).Checked;
  BApercu := TCheckBox(GetControl ('BAPERCU')).Checked;
  BDuplicata := TCheckBox(GetControl ('BDUPLICATA')).Checked;

	ImprViaTOB := nil;
  Pages:=Nil;

  OldQRThread := V_PGI.QRMultiThread;
  OldQR := V_PGI.QRPDF;
 	ChoixNonDirect := '';
  ModeleNonDirect := '';
  ChoixDirect := '';
  ModeleDirect := '';
  //
  ChoixAvancNonDirect := '';
  ModeleAvancNonDirect := '';
  OKEtatAvancNonDirect := false;
  //
  ChoixAvancDirect := '';
  ModeleAvancDirect := '';
  OKEtatAvancDirect := false;
  //
  OldprintDialog := V_PGI.NoPrintDialog;
  //
  //FV1 : 23/07/2015 - FS#1518 - ESPACS : Critère "Nombre exemplaires" ne sert pas ds Editions factures
  Nbexemplaires := ThSpinEdit(GetControl('NBEXEMPLAIRE')).Value;

  //
	for iInd := 0 to TobPiece.Detail.count - 1 do
  begin
    TOBP := TobPiece.Detail[IInd];
    //
    Avancement := pos(RenvoieTypeFact (TOBP.GetValue('GP_AFFAIREDEVIS')), 'AVA;DAC')>0;
    if ((Pos(NaturePieceG.Value,'FBT;DAC;FBP;BAC')>0)) and (Avancement) then
    begin
    	if BeditDirect then
      begin
        if (ChoixAvancDirect ='') then
        begin
          RenseignePiece (TOBP,BApercu,Choix,stModele,true,Nbexemplaires);
          ChoixAvancDirect := Choix;
          ModeleAvancDirect := stModele;
          OkEtatAvancDirect := TOBP.GetBoolean('OKETAT');
        end else
        begin
          SetInfoPieces (TOBP, ChoixAvancDirect + ModeleAvancDirect,OkEtatAvancDirect, NbExemplaires);
        end;
      end
      else
      begin
        if (ChoixAvancNonDirect ='') then
        begin
          RenseignePiece (TOBP,BApercu,Choix,stModele,false,Nbexemplaires);
          ChoixAvancNonDirect := Choix;
          ModeleAvancNonDirect := stModele;
          OKEtatAvancNonDirect := TOBP.GetBoolean('OKETAT');
        end else
        begin
          SetInfoPieces (TOBP, ChoixAvancNonDirect + ModeleAvancNonDirect,OKEtatAvancNonDirect, NbExemplaires);
        end;
      end;
    end
    else
    begin
    	if BeditDirect then
      begin
        if (ChoixDirect ='') then
        begin
          RenseignePiece (TOBP,BApercu,Choix,stModele,true,Nbexemplaires);
          ChoixDirect := Choix;
          ModeleDirect := stModele;
          OkEtatAvancDirect := TOBP.GetBoolean('OKETAT');
        end else
        begin
          SetInfoPieces ( TOBP, ChoixDirect + ModeleDirect, OkEtatAvancDirect, NbExemplaires);
        end;
      end else
      begin
        if (ChoixNonDirect ='') then
        begin
          RenseignePiece (TOBP,BApercu,Choix,stModele,false,Nbexemplaires);
          ChoixNonDirect := Choix;
          ModeleNonDirect := stModele;
          OKEtatAvancNonDirect  := TOBP.GetBoolean('OKETAT');
        end else
        begin
          SetInfoPieces (TOBP, ChoixNonDirect + ModeleNonDirect, OKEtatAvancNonDirect, NbExemplaires);
        end;
      end;
    end;
  end;

  if BEditDirect then
  begin
  	ImprViaTOB := TImprPieceViaTOB.create(ecran);
  end;
  //
  //
  TobPiece.Detail.Sort ('ETATSTANDARD_IMPMODELE;GP_NUMERO');
  StartPDFBatch;
  aborted := true;

  TRY
    for iInd := 0 to TobPiece.Detail.Count - 1 do
    begin
      TOBP := TobPiece.Detail[iInd];

      if (Not TOBP.GetValue ('OKETAT')) then Continue;

      //FV1 : 22/05/2015 - Contrôle pour interdire l'impression sur pièce visé et montant du visa si la nature est paramétrée en ce sens
      if Not ControleVisaBTP(TobPiece.detail[iInd]) then
      Begin
        StMsg := 'Impression pièce N° ' + IntToStr(TobPiece.detail[iInd].GetValue('GP_NUMERO')) + ' impossible. Pièce non visée !';
        PGIError(stMsg,'Gestion des Visas');
        Continue;
      end;

      Aborted := false;
      //
      BEditionDirect := BEditDirect;
      //
      //if (Pos(TobPiece.detail[iind].GetValue('GP_NATUREPIECEG'),'FBT;DAC;FBP;BAC')>0) then ///and (pos(RenvoieTypeFact (TobPiece.detail[iind].GetValue('GP_AFFAIREDEVIS')), 'AVA;DAC')>0) then
      //begin
      //  recapSituations := GetInfoRecapFromModele(TobPiece.detail[iind]);
      //end;

      if (Pos(TobP.getValue('GP_NATUREPIECEG'),'FBT;FBP')>0) then
      begin
        CalcReglementSituations (TobP.getValue('GP_AFFAIRE'));
      end;

      stModele  := copy (TobP.GetValue ('ETATSTANDARD_IMPMODELE'), 2, length (TobP.GetValue ('ETATSTANDARD_IMPMODELE')) - 1);
      //
      stCle     := TobP.GetValue ('GP_NATUREPIECEG')  + ';' +
      DateToStr   (TobP.GetValue ('GP_DATEPIECE'))    + ';' +
                   TobP.GetValue ('GP_SOUCHE')        + ';' +
      IntToStr    (TobP.GetValue ('GP_NUMERO'))       + ';' +
      IntToStr    (TobP.GetValue ('GP_INDICEG'));
      StringToCleDoc(stCle, CleDoc);

      if (not BEditionDirect) then
      begin
        {$IFDEF CCS3}
        if TobP.GetValue ('OKETAT') then
        begin
          stSql:=RecupSQLModele('E','GPJ', stModele,'','','',' WHERE '+WherePiece(CleDoc,ttdPiece,False));
          if (pos('ORDER BY',uppercase(stSql))<=0) then stSql:=stSql+' order by GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE,GL_ARTICLE' ;
          Pages := TPageControl.Create(Application);
        end
        else
        begin
          stSql:=RecupSQLModele('L','GPI', stModele,'','','',' WHERE '+WherePiece(CleDoc,ttdPiece,False));
          if (pos('ORDER BY',uppercase(stSql))<=0) then stSQL := stSQL+ ' ORDER BY GL_NUMLIGNE';
        end;
        {$ELSE}
        if TobP.GetValue ('OKETAT') then
        begin
          SQLOrder:='';
          stSql:=RecupSQLEtat('E','GPJ', stModele,'','','',' WHERE '+WherePiece(CleDoc,ttdPiece,False),SQLOrder);
          stSql:=stSql+' and GL_NONIMPRIMABLE<>"X"' ;
          if (SQLOrder='') then
            stSql:=stSql+' order by GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE,GL_ARTICLE'
          else
            stSql:=stSql+' '+SQLOrder;
          Pages := TPageControl.Create(Application);
        End
        else
        Begin
          SQLOrder:='';
          stSql:=RecupSQLEtat('L','GPI', stModele,'','','',' WHERE '+WherePiece(CleDoc,ttdPiece,False),SQLOrder);
          if (SQLOrder='') then
            stSql := stSql+ ' ORDER BY GL_NUMLIGNE'
          else
            stSql:=stSql+' '+SQLOrder;
        End;
        {$ENDIF}
      end
      else
      begin
        ImprViaTOB.ClearInternal;
        ImprViaTOB.SetDocument (cledoc,false);
        ImprViaTOB.PreparationEtat;
      end;

      iNbPiece := StrToInt(TobPiece.Detail[iInd].GetValue ('NBEXEMPLAIRE'));
      if iNbPiece = 0 then iNbPiece :=1;

      for iNbExemplaire := 1 to iNbPiece do
      begin
        TL := TList.Create ;
        TT := TStringList.Create ;
        TT.Add (stSql);
        TL.Add (TT);
        if TobP.GetValue ('OKETAT') then
        begin
          if (not BEditionDirect) then
            EditionNonDirecte(TOBP,CleDoc, stModele, stSql, BDuplicata)
          else
            EditionDirecte(TOBP, stModele, ImprViaTOB, BDuplicata, Avancement);
        end;
        BAp:=false;
        TT.Free;
        TL.Free;
      end;


        //if TobPiece.Detail[iInd].GetValue ('OKETAT') then
        //BEGIN
          {*
          if (not BEditionDirect) then
          begin
            if LanceEtat('E','GPJ',stModele,TRUE,False,False,Nil,trim (stSql),'',BDuplicata)<0 then
            	V_PGI.IoError:=oeUnknown ;
          end
          else
          begin
            if (stModele <> '') or Not (RecapSituations) then
            begin
            TheModele:=stModele;
            If TheModele = '' Then TheModele := ImprViaTOB.GetModeleAssocie(stModele);
            UneTOB := ImprViaTOB.TOBAIMPRIMER;
              if LanceEtatTOB ('E','GPJ',TheModele,UneTOB,TRUE,false,false,nil,'','',Bduplicata) < 0 then V_PGI.IoError:=oeUnknown ;
          end;
          end;
          *}
          // Edition du recap de situations si besoin est
          //if (V_PGI.IoError = OeOk) and (RecapSituations) then
          //begin
            //if (not BEditionDirect) then
            //begin
            //  SQLOrder:='';
            //  ModeleR:=GetParamsoc('SO_BTETATSIR');
            //  if ModeleR='' then ModeleR:=Copy(stModele,1,2)+'R';
            //  stSQL:=RecupSQLEtat('E','GPJ',ModeleR,'','','',' WHERE '+WherePiece(CleDoc,ttdPiece,False),SQLOrder) ;
            //  stSQL:=stSql+' '+SQLOrder;
            //  if LanceEtat('E','GPJ',ModeleR,TRUE,False,False,Nil,Trim(stSQL),'',Bduplicata) < 0 then
            //    V_PGI.ioError := oeUnknown;
            //end
            //else
            //begin
            //  if (Pos(TobPiece.detail[iind].GetValue('GP_NATUREPIECEG'),'FBP;BAC;')>0) then
            //  begin
            //    ModeleR:=GetParamsoc('SO_BTETATFBR');
            //  end else
            //  begin
            //    if Avancement then
            //      ModeleR:=GetParamsoc('SO_BTETATSIRDIR')
            //    else
            //      ModeleR:=GetParamsoc('SO_BTETATRECAPDIRECT');
            //  end;
            //  if ModeleR <> '' then
            //  begin
            //    UneTOB := ImprViaTOB.TOBRECAP;
            //    if LanceEtatTOB ('E','GPJ',ModeleR,UneTOB,TRUE,false,false,nil,'','',Bduplicata) < 0 then V_PGI.IoError:=oeUnknown ;
            //  end;
            //end;
          //end;
          //Edition du Récap de sous-traitance
          //FV1 : 14/12/2016 - FS#2280 - DELABOUDINIERE - MULTIPHONE NETCOM - editions différées des contrats
          //if BEditDirect then
          //begin
          //  if ImprViaTOB.TobSsTraitantImp.Detail.count <> 0 then
          //  begin
          //    if ImprViaTOB.TobSsTraitantImp.Detail[0].detail.count > 5 then
          //    begin
          //      ModeleR:=GetParamsoc('SO_BTETATSSTRAITANCE');
          //      if ModeleR <> '' then
          //      begin
          //        UneTOB := ImprViaTOB.TobSsTraitantImp;
          //        if LanceEtatTOB ('E','GPJ',ModeleR,UneTOB,TRUE,false,false,nil,'','',Bduplicata) < 0 then V_PGI.IoError:=oeUnknown ;
          //      end;
          //    end;
          //  end;
          //Edition du Recap des Retenues diverses
          (*
          if ImprViaTOB.TobRetenueDiverse.Detail.count <> 0 then
          begin
            if ImprViaTOB.TobRetenueDiverse.Detail[0].detail.count > 5 then
            begin
              ModeleR:=GetParamsoc('SO_BTETATSSTRAITANCE');
              if ModeleR = '' then break;
              UneTOB := ImprViaTOB.TobRetenueDiverse;
              if LanceEtatTOB ('E','GPJ',ModeleR,UneTOB,TRUE,false,false,nil,'','',Bduplicata) < 0 then V_PGI.IoError:=oeUnknown ;
            end;
          end;
          *)
          //end;
          // ---
        //END;
        //BAp:=false;
        //TT.Free;
        //TL.Free;
      //end;

      if TobP.GetValue ('OKETAT') then Pages.Free;

      if GetControlText ('GP_EDITEE') = '-' then TOBP.PutValue ('GP_EDITEE', 'X');

      if BEditDirect then ImprViaTOB.ClearInternal;

      if V_PGI.ioError <> oeOk then break;
    end;
  FINALLY
		CancelPDFBatch;
  END;

{$IFNDEF EAGLCLIENT}
	if not aborted then
  begin
    if Bapercu then
    begin
      V_PGI.NbDocCopies := 1; // puisque le nombre d'exemplaire de chaque état est déja traité dans la préparation
      PreviewPDFFile('',GetMultiPDFPath,True);
    end else
    begin
      PrintD := TPrintDialog.Create(ecran);
      printD.Copies := 1;
      if printD.Execute then
      begin
        PrintPdf(GetMultiPDFPath, Printer.Printers[Printer.PrinterIndex], '');
      end;
      PrintD.free;
    end;
  end;
{$ENDIF}

  if BEditDirect then ImprViaTOB.free;

  V_PGI.QRPDF := OldQR;
  V_PGI.QRPDFQueue := '';
  V_PGI.QRPDFMerge := '';
  V_PGI.QRMultiThread := OldQRThread;
  V_PGI.PrintToFile := false;
  V_PGI.QRPrinted := false;
  v_pgi.MiseSousPli := False;
  v_pgi.PageMiseSousPli := 0;

end;


Procedure TOF_BTEditDocDiff_Mul.EditionNonDirecte(TOBP : TOB; Cledoc : R_CleDoc; StModele : string; StSQl : string; BDuplicata : boolean);
Var SQLOrder        : String;
    BRecapsituation : Boolean;
    TheModele       : string;
begin

  BRecapsituation := false;
  TheModele := StModele;
  //
  if (Pos(TOBP.GetValue('GP_NATUREPIECEG'),'FBT;DAC;FBP;BAC')>0) then BRecapsituation := GetInfoRecapFromModele(TOBP);

  if LanceEtat('E','GPJ',TheModele,TRUE,False,False,Nil,trim (stSql),'',BDuplicata) < 0 then V_PGI.IoError:=oeUnknown ;

  if (V_PGI.IoError = OeOk) and (BRecapsituation) then
  begin
    SQLOrder  :='';
    TheModele   := GetParamsocSecur('SO_BTETATSIR', '');
    if TheModele = '' then TheModele := Copy(stModele,1,2) + 'R';
    stSQL     :=RecupSQLEtat('E','GPJ',TheModele,'','','',' WHERE ' + WherePiece(CleDoc,ttdPiece,False),SQLOrder) ;
    stSQL     :=stSql + ' ' + SQLOrder;
    if LanceEtat('E','GPJ',TheModele,TRUE,False,False,Nil,Trim(stSQL),'',Bduplicata) < 0 then V_PGI.ioError := oeUnknown;
  end;

end;

Procedure TOF_BTEditDocDiff_Mul.EditionDirecte(TOBP : TOB; StModele : string; ImprViaTOB : TImprPieceViaTOB; BDuplicata, Avancement : boolean);
Var BRecapsituation : Boolean;
    TheModele       : string;
    UneTob          : TOB;
begin

  BRecapsituation := false;
  //
  if (Pos(TOBP.GetValue('GP_NATUREPIECEG'),'FBT;DAC;FBP;BAC')>0) then BRecapsituation := GetInfoRecapFromModele(TOBP);

  TheModele := StModele;

  if (TheModele <> '') or Not (BRecapsituation) then
  begin
    If TheModele = '' Then TheModele := ImprViaTOB.GetModeleAssocie(TheModele);
    UneTOB      := ImprViaTOB.TobAImprimer;
    if LanceEtatTOB ('E', 'GPJ', TheModele, UneTob, TRUE, false, false, nil, '', '', Bduplicata) < 0 then V_PGI.IoError:=oeUnknown ;
  end;

  if (Pos(TOBP.GetValue('GP_NATUREPIECEG'),'FBP;BAC;')>0) then
    TheModele   := GetParamsocSecur('SO_BTETATFBR', '')
  else
  begin
    if Avancement then
      TheModele := GetParamsocSecur('SO_BTETATSIRDIR', '')
    else
      TheModele := GetParamsocSecur('SO_BTETATRECAPDIRECT', '');
  end;

  if (V_PGI.IoError = OeOk) and (BRecapsituation) then
  begin
    if TheModele <> '' then
    begin
      UneTOB      := ImprViaTOB.TOBRECAP;
      if LanceEtatTOB ('E','GPJ',TheModele,UneTOB,TRUE,false,false,nil,'','',Bduplicata) < 0 then V_PGI.IoError:=oeUnknown ;
    end;
  end;

  if ImprViaTOB.TobSsTraitantImp.Detail.count <> 0 then
  begin
    if ImprViaTOB.TobSsTraitantImp.Detail[0].detail.count > 5 then
    begin
      TheModele   := GetParamsocSecur('SO_BTETATSSTRAITANCE', '');
      if TheModele <> '' then
      begin
        UneTOB    := ImprViaTOB.TobSsTraitantImp;
        if LanceEtatTOB ('E', 'GPJ', TheModele, UneTOB, TRUE, false, false, nil, '', '', Bduplicata) < 0 then V_PGI.IoError:=oeUnknown ;
      end;
    end;
  end;

  //Edition du Recap des Retenues diverses
  if ImprViaTOB.TobRetenueDiverse.Detail.count <> 0 then
  begin
    if ImprViaTOB.TobRetenueDiverse.Detail[0].detail.count > 5 then
    begin
      TheModele :=GetParamsocSecur('SO_BTETATSSTRAITANCE', '');
      if TheModele = '' then exit;
      UneTOB    := ImprViaTOB.TobRetenueDiverse;
      if LanceEtatTOB ('E','GPJ',TheModele,UneTOB,TRUE,false,false,nil,'','',Bduplicata) < 0 then V_PGI.IoError:=oeUnknown ;
    end;
  end;

end;


procedure TOF_BTEditDocDiff_Mul.ModificationModeleBTP(var OkEtat, BApercu: boolean; var Modele : string; var NbCopies : integer; EditDirect : boolean; LibelleAff :string);
var NouvelleValeur : string;
		Apercu,modeleEtat,IsEtat,stTitre : string;
begin
	if EditDirect then stTitre := 'Sélection du modèle '+LibelleAff+' direct'
  							else stTitre := 'Sélection du modèle '+LibelleAff+' non direct';
  if BApercu then Apercu:='X' else Apercu:='-' ;
  ModeleEtat:=Modele ;
  NouvelleValeur:=AGLLanceFiche('MBO','VALIDMODELE','','','MODELEETAT='+ModeleEtat+';'
                  +';MODELEETIQ=;APERCU='+Apercu+';APERCUETIQ=-;IMPETIQ=-;TITRE='+stTitre+';'+
                  'NBCOPIE='+IntToStr(NbCopies)+'') ;
  if NouvelleValeur <>'' then
  begin
    IsEtat:=ReadTokenSt(NouvelleValeur) ;
    if IsEtat='X' then OkEtat:=True else OkEtat:=False ;
    Modele:=ReadTokenSt(NouvelleValeur) ;
    Apercu:=ReadTokenSt(NouvelleValeur) ;
    if Apercu='X' then Bapercu:=True else BApercu:=False ;
    NBCopies:=StrToInt(ReadTokenST(NouvelleValeur)) ;
  end else OkEtat:=False ;
end;

procedure TOF_BTEditDocDiff_Mul.RenseignePiece (TobP : TOB; var BApercu : boolean; var Choix : string; var Modele: string; EditDirect : boolean; var Nbcopies : integer);
var
		QTiers : TQuery;
    stEtatStandard, Suff : string;
    NaturePiece, {Modele,} Domaine, Etab :string;
    OkEtat: boolean;
    LibellePiece : string;
begin
  if (not TOBP.FieldExists ('ETATSTANDARD_IMPMODELE')) then
  begin
    TobP.AddChampSup ('ETATSTANDARD_IMPMODELE', True);
    TobP.AddChampSup ('NBEXEMPLAIRE', true);
    TobP.AddChampSup ('OKETAT', True);
  end;

  NaturePiece:=TobP.GetValue('GP_NATUREPIECEG') ;
  OkEtat:=False ;
  LibellePiece := rechDom ('GCNATUREPIECEG',NaturePiece,false);
  Etab:=TobP.GetValue('GP_ETABLISSEMENT') ; Domaine:=TobP.GetValue('GP_DOMAINE') ;
  Suff:='' ; if TobP.GetValue('GP_SAISIECONTRE')='X' then Suff:='CON' ;
  Modele:='' ;
  if NbCopies = 0 then
  begin
    NbCopies:=GetInfoParPiece(NaturePiece,'GPP_NBEXEMPLAIRE') ; if ((NbCopies<0) or (NbCopies>99)) then NbCopies:=0 ;
  end;

  {Recherche sur profils tiers}
  QTiers := OpenSQL('SELECT GTP_IMPMODELE' + Suff + ', GTP_IMPETAT' + Suff
         + ', GTP_NBEXEMPLAIRE FROM TIERSPIECE WHERE GTP_TIERS="' + TobP.GetValue ('GP_TIERS')
         + '" AND GTP_NATUREPIECEG="' + TobP.GetValue ('GP_NATUREPIECEG') + '"',false);
  if not QTiers.Eof then
     BEGIN
     Modele:=QTiers.FindField('GTP_IMPETAT'+Suff).AsString ;
     if Modele<>'' then OkEtat:=True else Modele:=QTiers.FindField('GTP_IMPMODELE'+Suff).AsString ;
     NbCopies:=QTiers.FindField('GTP_NBEXEMPLAIRE').AsInteger ; if ((NbCopies<0) or (NbCopies>99)) then NbCopies:=0 ;
     stEtatStandard := '4';
     END ;
  Ferme(QTiers);

  {Recherche sur "les "PARPIECE}
  if Modele='' then
     BEGIN
     {Etat Nature/Domaine}
     stEtatStandard := '3' ;
     Modele:=GetInfoParPieceDomaine(NaturePiece,Domaine,'GDP_IMPETAT'+Suff) ;
     if Modele<>'' then  OkEtat:=True else
        BEGIN
        {Document Nature/Domaine}
        Modele:=GetInfoParPieceDomaine(NaturePiece,Domaine,'GDP_IMPMODELE'+Suff) ;
        if Modele='' then
           BEGIN
           stEtatStandard := '2' ;
           {Etat Nature/Etablissement}
           Modele:=GetInfoParPieceCompl(NaturePiece,Etab,'GPC_IMPETAT'+Suff) ;
           if Modele<>'' then OkEtat:=True else
              BEGIN
              {Document Nature/Etablissement}
              Modele:=GetInfoParPieceCompl(NaturePiece,Etab,'GPC_IMPMODELE'+Suff) ;
              if Modele='' then
                 BEGIN
                 stEtatStandard := '1' ;
                 if EditDirect then
                 begin
                   Modele:=GetInfoParPiece(NaturePiece, 'GPP_IMPMODELE');
                   if Modele <> '' then OkEtat := true;
                 end else
                 begin
                  {Etat Nature}
                   Modele:=GetInfoParPiece(NaturePiece,'GPP_IMPETAT'+Suff) ;
                  {Modèle Nature}
                   if Modele<>'' then OkEtat:=True else Modele:=GetInfoParPiece(NaturePiece,'GPP_IMPMODELE'+Suff) ;
                 end;
                 END ;
              END ;
           END ;
        END ;
     END ;

  // Recherche sur Paramsoc
  if Modele='' then
  begin
    stEtatStandard := '0' ;
    Modele:=VH_GC.GCImpModeleDefaut ;
  end;

  // --- Modif BTP
  // Recherche spéciale pour le modèle des situations de travaux (BTP)
  {$IFDEF BTP}
    if (Pos(NaturePiece,'FBT;DAC;FBP;BAC')>0) and (pos(RenvoieTypeFact (TobP.GetValue('GP_AFFAIREDEVIS')), 'AVA;DAC')>0) then
    begin
      LibellePiece := 'Situation';
      if EditDirect then
      begin
        // Récupérer le modèle des situations dans les paramètres
        if Pos(NaturePiece,'DAC;FBT') > 0 then  Modele := GetParamsoc('SO_BTETATSITDIR')
                                          else  Modele := GetParamsoc('SO_BTETATFBP');
      end else
      begin
        // Récupérer le modèle des situations dans les paramètres
        Modele := GetParamsoc('SO_BTETATSIT');
      end;
      OkEtat:=True;
    end else
    begin
      if Modele <> '' then OkEtat := true;
    end;

    if OkEtat then
    begin
      if GetInfoParPiece(NaturePiece, 'GPP_VALMODELE') = 'X' then
      begin
        ModificationModeleBTP(OkEtat, BApercu, Modele,NbCopies,EditDirect,LibellePiece);
      end;
    end;
    if Bapercu then TCheckBox(GetControl ('BAPERCU')).Checked := true
               else TCheckBox(GetControl ('BAPERCU')).Checked := false;
  {$ENDIF}

  Choix := stEtatStandard;
  TobP.PutValue ('ETATSTANDARD_IMPMODELE', stEtatStandard + Modele);
  TobP.PutValue ('OKETAT', OkEtat);
  TobP.PutValue ('NBEXEMPLAIRE', NbCopies);
  if GetInfoParPiece(NaturePiece,'GPP_ACTIONFINI') = 'IMP' then TobP.PutValue ('GP_VIVANTE', '-');
end;



procedure TOF_BTEditDocDiff_Mul.SetInfoPieces (TobP : TOB; ETATSTANDARD : string; OKETAT : boolean; NbExemplaire : integer);
var naturePiece : string;
begin
	NaturePiece:=TobP.GetValue('GP_NATUREPIECEG') ;
  if (not TOBP.fieldExists('ETATSTANDARD_IMPMODELE')) then
  begin
    TobP.AddChampSup ('ETATSTANDARD_IMPMODELE', True);
    TobP.AddChampSup ('NBEXEMPLAIRE', true);
    TobP.AddChampSup ('OKETAT', True);
  end;
  TobP.PutValue ('ETATSTANDARD_IMPMODELE', ETATSTANDARD);
  TobP.PutValue ('OKETAT', OKETAT);
  TobP.PutValue ('NBEXEMPLAIRE', NbExemplaire);
  if GetInfoParPiece(NaturePiece,'GPP_ACTIONFINI') = 'IMP' then TobP.PutValue ('GP_VIVANTE', '-');
end;

/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLEditDocDiff_mul(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_BTEditDocDiff_Mul) then TOF_BTEditDocDiff_Mul(MaTOF).LanceEdition else exit;
end;

procedure TOF_BTEditDocDiff_Mul.OnLoad;
begin
  inherited;
(*
	if NATUREPIECEG.value = '' then
  begin
  	PGIInfo('Attention : La nature de pièce est marquée comme masquée#13#10Vous devez la définir comme non masquée');
    PostMessage(TFmul(Ecran).Handle, WM_CLOSE, 0, 0);
    Exit;
  end;
*)
end;

Initialization
registerclasses([TOF_BTEditDocDiff_Mul]);
RegisterAglProc('BTEditDocDiff_mul',TRUE,1,AGLEditDocDiff_mul);
end.
