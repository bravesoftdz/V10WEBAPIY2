unit FermeturePieces;
 
interface
Uses Classes,
     Windows,
     sysutils,
     ComCtrls,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
{$ELSE}
     MainEagl,
{$ENDIF}
     HCtrls,
     Hpanel,
     HEnt1,
     HMsgBox,
     Vierge,
     EntGc,
     BTPUtil,
     ExtCtrls,
     SaisUtil,
     UtilPgi,
     UtilFermeturePiece,
     UTOB;

type

	TFermeturePiece = class (TOBject)
  	private
    	fTOBpiece : TOB;
      ffermeture : boolean;
      fresult : integer;
      function controlePieceOk : boolean;
      procedure TraitePiece;
    	function controleDevis: integer;
    	function ControleFacture: integer;
    	function ControleAvoir: integer;
    	function ControleChantierInUse: integer;
    public
      property Lapiece : TOB read fTOBPiece write fTOBPiece;
      property Fermeture : boolean read ffermeture write ffermeture;
      property result : integer read fresult;
      procedure lancetraitement;
  end;

function  TraitementFermeturePiece (TOBPiece : TOB; Mode : boolean) : integer;
implementation

uses StrUtils;

function  TraitementFermeturePiece (TOBPiece : TOB; Mode : boolean) : integer;
var FermeturePiece : TFermeturePiece;
begin
   FermeturePiece := TFermeturePiece.create;
   FermeturePiece.LaPiece := TOBPiece;
   FermeturePiece.Fermeture := Mode;
   TRY
     FermeturePiece.lanceTraitement;
   FINALLY
     result := FermeturePiece.Result;
   	 FermeturePiece.free;
   END;
end;

{ TFermeturePiece }

function TFermeturePiece.ControleDevis : integer;
begin
	result := 0;
  if (not ffermeture) then
  begin
  	// on veut reouvrir un devis
    if fTOBPiece.GetValue('GP_AFFAIREDEVIS')<> '' then
    begin
      if (ExisteSql('SELECT AFF_AFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE="'+
                       fTOBPiece.GetValue('GP_AFFAIREDEVIS')+
                       '" AND AFF_ETATAFFAIRE="TER"')) then
      begin
        result := 1;
        exit;
      end;
    end;
  end;
end;

function TFermeturePiece.ControleFacture : integer;
var QQ : TQuery;
		TheNum : integer;
    Req : string;
begin
	result := 0;
  if (not ffermeture) then
  begin
  	// on veut reouvrir une Facture
    if fTOBPiece.GetValue('GP_AFFAIREDEVIS')<> '' then
    begin
    	// verif dossier de facturation non fermé
      if (ExisteSql('SELECT AFF_AFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE="'+
                       fTOBPiece.GetValue('GP_AFFAIREDEVIS')+
                       '" AND AFF_ETATAFFAIRE="TER"')) then
      begin
        result := 1;
        exit;
      end;
      if fTOBPiece.getVAlue('AFF_GENERAUTO')='AVA' then
      begin
        //------------
        {Dans le cas de facture d'avancement}
        // verif si derniere facture du dossier de facturation
        REq := 'SELECT B2.BST_NUMEROFAC FROM BSITUATIONS B2 WHERE '+
                       'B2.BST_NUMEROSIT= (SELECT MAX(B1.BST_NUMEROSIT) FROM BSITUATIONS B1 WHERE '+
                       'B1.BST_SSAFFAIRE="'+fTOBPiece.GetVAlue('GP_AFFAIREDEVIS')+'") AND B2.BST_SSAFFAIRE='+
                       '"'+fTOBPiece.GetVAlue('GP_AFFAIREDEVIS')+'"';
        QQ := OPeNSql (req,true,1,'',true);
        if not QQ.eof then
        begin
          TheNum := QQ.FindField('BST_NUMEROFAC').AsInteger;
        end else
        begin
          TheNum := fTOBPiece.getValue('GP_NUMERO');
        end;
        ferme (QQ);
        if TheNUm <> fTOBPiece.getValue('GP_NUMERO') then
        begin
          result := 11;
          exit;
        end;
      end;
    end;
    //
  end;
end;

function TFermeturePiece.ControleChantierInUse : integer;
begin
	result := 0;
  if (not ffermeture) then
  begin
  	// on veut reouvrir une piece
    if fTOBPiece.GetValue('GP_AFFAIRE')<> '' then
    begin
      if (ExisteSql('SELECT AFF_AFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE="'+
                       fTOBPiece.GetValue('GP_AFFAIRE')+
                       '" AND AFF_ETATAFFAIRE="TER"')) then
      begin
        result := 20;
        exit;
      end;
    end;
  end;
end;

function TFermeturePiece.ControleAvoir : integer;
begin
	result := 0;
  if (not ffermeture) then
  begin
  	// on veut reouvrir un Avoir
    if fTOBPiece.GetValue('GP_AFFAIREDEVIS')<> '' then
    begin
    	// verif dossier de facturation non fermé
      if (ExisteSql('SELECT AFF_AFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE="'+
                       fTOBPiece.GetValue('GP_AFFAIREDEVIS')+
                       '" AND AFF_ETATAFFAIRE="TER"')) then
      begin
        result := 1;
        exit;
      end;
    end;
    //
  end;
end;

function TFermeturePiece.ControlePieceOk: boolean;
var NaturePiece : string;
begin
	NaturePiece := fTOBPiece.getValue('GP_NATUREPIECEG');
  if NaturePiece='DBT' then
  begin
    fresult := ControleDevis;
  end else if Pos(NaturePiece,'FBT;DAC;FBP;BAC')>0 then
  begin
    fresult := ControleFacture;
  end else if (NaturePiece='ABT') or (NaturePiece='ABP') then
  begin
    fresult := ControleAvoir;
  end else
  begin
    fresult := ControleChantierInUse;
  end;
  //
  result := (fresult= 0);
end;

procedure TFermeturePiece.lancetraitement;
var io   : TIoErr ;
begin
	if controlePieceOk then
  begin
    Io := Transactions (TraitePiece, 0);
    Case Io of
      oeUnKnown :BEGIN
                  fResult := -1;
                 END;
    end;
  end;
end;

procedure TFermeturePiece.TraitePiece;
var VIVANTE : String;
    sQteMaj : String;
    StSQL   : String;
begin
  if ffermeture then VIVANTE := '-' else VIVANTE := 'X';
  if ffermeture then reajusteStock (fTOBPiece); // uniquement si l'on veut fermer la piece
  if V_PGI.Ioerror = OeOk then
  begin
    If ExecuteSql ('UPDATE PIECE SET GP_VIVANTE="' + VIVANTE +
                '" WHERE '+
                'GP_NATUREPIECEG="'+fTOBPiece.getValue('GP_NATUREPIECEG')+'" '+
                'AND GP_SOUCHE="'+fTOBPiece.getValue('GP_SOUCHE')+'" '+
                'AND GP_NUMERO='+InttoStr(fTOBPiece.getValue('GP_NUMERO'))+' '+
                'AND GP_INDICEG='+IntToStr(fTOBPiece.getValue('GP_INDICEG'))) < 0 then V_PGI.IoError := OeUnknown;
  end;
  if V_PGI.Ioerror = OeOk then
  begin
  	if (not ffermeture) and (pos(fTOBPiece.getValue('GP_NATUREPIECEG'),'B00;DBT;FBT;ABT;ABP;FF;AF;AFS')>0) then
    begin
    	sQteMaj := ',GL_QTERESTE=GL_QTEFACT, GL_QTESTOCK=GL_QTEFACT, GL_MTRESTE=GL_MONTANTHTDEV ';
    end
    else if (not ffermeture) then
    begin
    	sQteMaj := '';
    end else begin
    	sQteMaj := ',GL_QTERESTE=0, GL_QTESTOCK=0, GL_MTRESTE=0 ';
    end;
    If ExecuteSql ('UPDATE LIGNE SET GL_VIVANTE="'+VIVANTE+'" '+ sQteMaj+
    						'WHERE '+
                'GL_NATUREPIECEG="'+fTOBPiece.getValue('GP_NATUREPIECEG')+'" '+
                'AND GL_SOUCHE="'+fTOBPiece.getValue('GP_SOUCHE')+'" '+
                'AND GL_NUMERO='+InttoStr(fTOBPiece.getValue('GP_NUMERO'))+' '+
                'AND GL_INDICEG='+IntToStr(fTOBPiece.getValue('GP_INDICEG'))) < 0 then V_PGI.IoError := OeUnknown;
    //FV1 - 20/06/2016 - FS#2003 - MATFOR - devis refusé ne se met pas ENCOURS une fois le devis réouvert
    if (not ffermeture) then
    begin
      StSQL := 'UPDATE AFFAIRE SET AFF_ETATAFFAIRE="ENC" ';
      StSQL := StSQL + 'FROM AFFAIRE LEFT JOIN PIECE ON AFF_AFFAIRE=GP_AFFAIREDEVIS ';
      StSQL := StSQL + 'WHERE AFF_ETATAFFAIRE="REF" ';
      StSQL := StSQL + 'AND GP_NATUREPIECEG="' + fTOBPiece.getValue('GP_NATUREPIECEG')     +'" ';
      StSQL := StSQL + 'AND GP_SOUCHE="'       + fTOBPiece.getValue('GP_SOUCHE')           +'" ';
      StSQL := StSQL + 'AND GP_NUMERO='        + InttoStr(fTOBPiece.getValue('GP_NUMERO')) +' ';
      StSQL := StSQL + 'AND GP_INDICEG='       + IntToStr(fTOBPiece.getValue('GP_INDICEG'));
      if ExecuteSQL(StSQL) < 0 then V_PGI.IoError := OeUnknown;
    end;
  end;
end;

end.
