unit UtilDuplicBordereaux;

interface
Uses StdCtrls, 
     Controls, 
     Classes,
     M3FP, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     fe_main,
{$else}
     eMul,
     uTob,
     MaineAGL,
{$ENDIF}
		 UTOB,	
		 utofAfBaseCodeAffaire,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     EntGc,
     FactComm,
     UTOF,uEntCommun ;

Type
  TOF_BTSELBORDEREAUX = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
    procedure DupliqueBordereau (Piece,Libelle : string);
  end ;

  TOF_BORDEREAU_AFF= Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
    private
    	CledocOrigine : r_cledoc;
    function ControleSaisie: boolean;
  end ;

procedure DuplicBordereau;
procedure AglDupliqueBordereaux ( parms: array of variant; nb: integer ) ;
procedure EcritEnteteBord(X : TForm;TOBPiece : TOB);

implementation
uses facture,Factutil;

procedure DuplicBordereau;
begin
	AglLanceFiche ('BTP','BTSELBORDEREAUX','','','ACTION=MODIFICATION');
end;

procedure TOF_BTSELBORDEREAUX.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSELBORDEREAUX.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSELBORDEREAUX.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTSELBORDEREAUX.OnLoad ;
begin
  Inherited ;
  THEdit(GetControl('BTBAFFAIRE0')).Enabled := false;
  THEdit(GetControl('BTBAFFAIRE1')).Enabled := false;
  THEdit(GetControl('BTBAFFAIRE2')).Enabled := false;
  THEdit(GetControl('BTBAFFAIRE3')).Enabled := false;
  THEdit(GetControl('BTBAVENANT')).Enabled := false;
end ;

procedure TOF_BTSELBORDEREAUX.OnArgument (S : String ) ;
begin
	fMulDeTraitement := true;
  Inherited ;
  THEdit(getControl('XX_WHERE')).Text :=' AND BDE_QUALIFNAT="PRINC1" AND BDE_PIECEASSOCIEE<>""';
  THEdit(GetControl('BTBAFFAIRE0')).Enabled := false;
  THEdit(GetControl('BTBAFFAIRE1')).Enabled := false;
  THEdit(GetControl('BTBAFFAIRE2')).Enabled := false;
  THEdit(GetControl('BTBAFFAIRE3')).Enabled := false;
  THEdit(GetControl('BTBAVENANT')).Enabled := false;
end ;

procedure TOF_BTSELBORDEREAUX.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTSELBORDEREAUX.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSELBORDEREAUX.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTSELBORDEREAUX.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2,
  Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin
  inherited;
  Aff:=THEdit(GetControl('BDE_AFFAIRE'));
  aff0 := THEdit(GEtControl('BTBAFFAIRE0'));
  Aff1:=THEdit(GetControl('BTBAFFAIRE1')); Aff2:=THEdit(GetControl('BTBAFFAIRE2'));
  Aff3:=THEdit(GetControl('BTBAFFAIRE3')); Aff4:=THEdit(GetControl('BTBAVENANT'));
  Tiers:=THEdit(GetControl('BDE_CLIENT'));
end;

procedure AglDupliqueBordereaux ( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     LaTof : TOF;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFMul) then Latof:=TFMul(F).Latof else exit;
  if (LaTof is TOF_BTSELBORDEREAUX) then TOF_BTSELBORDEREAUX(LaTof).DupliqueBordereau (Parms[1],Parms[2]) else exit;
end;

procedure TOF_BTSELBORDEREAUX.DupliqueBordereau(Piece,Libelle: string);
var Cledoc : r_cledoc;
begin
	DecodeRefPiece (Piece,Cledoc);
	AGLLanceFiche ('BTP','BTBORDEREAU_AFF','','','NATURE='+Cledoc.NaturePiece+';'+
  																						 'SOUCHE='+Cledoc.Souche+';'+
                                               'NUMERO='+InttoStr(cledoc.NumeroPiece)+';'+
                                               'INDICE='+InttoStr(cledoc.Indice)+';'+
                                               'LIBELLE='+Libelle);
end;

{ TOF_BORDEREAU_AFF }

procedure TOF_BORDEREAU_AFF.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2,
  Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin
  inherited;
  Aff:=THEdit(GetControl('BDE_AFFAIRE'));
  aff0 := THEdit(GEtControl('BTBAFFAIRE0'));
  Aff1:=THEdit(GetControl('BTBAFFAIRE1')); Aff2:=THEdit(GetControl('BTBAFFAIRE2'));
  Aff3:=THEdit(GetControl('BTBAFFAIRE3')); Aff4:=THEdit(GetControl('BTBAVENANT'));
  Tiers:=THEdit(GetControl('GP_TIERS'));
end;

procedure TOF_BORDEREAU_AFF.OnArgument(S: String);
var Arguments,critere,ChampMul,ValMul : string;
		X : integer;
begin
	fMulDeTraitement := true;
  inherited;
  Arguments := S;
  repeat
    Critere := uppercase(Trim(ReadTokenSt(Arguments)));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));

        if ChampMul = 'NATURE' then
        begin
          CledocOrigine.NaturePiece := ValMul;
        end else if ChampMul = 'SOUCHE' then
        begin
          CledocOrigine.Souche := ValMul;
        end else if ChampMul = 'NUMERO' then
        begin
          CledocOrigine.NumeroPiece := StrToInt(ValMul);
        end else if ChampMul = 'INDICE' then
        begin
          CledocOrigine.Indice := StrToInt(ValMul);
        end else if ChampMul = 'LIBELLE' then
        begin
          THEdit(GetControl('BDE_DESIGNATION')).text := ValMul;
        end;
      end;
    end;
  until Critere = '';
end;

procedure TOF_BORDEREAU_AFF.OnCancel;
begin
  inherited;

end;

procedure TOF_BORDEREAU_AFF.OnClose;
begin
  inherited;

end;

procedure TOF_BORDEREAU_AFF.OnDelete;
begin
  inherited;

end;

procedure TOF_BORDEREAU_AFF.OnDisplay;
begin
  inherited;

end;

procedure TOF_BORDEREAU_AFF.OnLoad;
begin
  inherited;

end;

procedure TOF_BORDEREAU_AFF.OnNew;
begin
  inherited;

end;

procedure TOF_BORDEREAU_AFF.OnUpdate;
begin
  inherited;
  if not ControleSaisie then begin lasterror := 1; exit end;
  DupliqueBordereauBTP (CledocOrigine,
  											THEdit(GetControl('GP_TIERS')).Text,
                        THEdit(GetControl('BDE_AFFAIRE')).Text,
                        THEdit(GetControl('BDE_DESIGNATION')).Text,
                        THEdit(getControl('BDE_DATE')).Text,
                        THEdit(getControl('BDE_DATE_')).Text);
end;

function TOF_BORDEREAU_AFF.ControleSaisie : boolean;

	function ClientExist (UnClient : string) : boolean;
  var QQ : TQuery;
  begin
  	QQ := OpenSql ('SELECT T_LIBELLE FROM TIERS WHERE T_TIERS="'+UnCLient+'" AND T_NATUREAUXI="CLI"',true);
    result := (not QQ.eof);
    ferme (QQ);
  end;

	function AffaireExist (UneAffaire : string) : boolean;
  var QQ : TQuery;
  begin
  	QQ := OpenSql ('SELECT AFF_LIBELLE FROM AFFAIRE WHERE AFF_AFFAIRE="'+UneAffaire+'"',true);
    result := (not QQ.eof);
    ferme (QQ);
  end;

var TheClient,TheAffaire : string;
begin
	result := true;
	TheClient := THEdit(GetControl('GP_TIERS')).Text;
  TheAffaire := THEdit(GetControl('BDE_AFFAIRE')).Text;
	if (TheClient = '' ) and (TheAffaire = '' ) then begin result := false;TForm(Ecran).ModalResult:=0;exit; end;
	if (TheClient <> '' ) and (not ClientExist (TheClient)) then
  begin
  	PgiBox ('Ce client n''existe pas');
  	result := false;
    TForm(Ecran).ModalResult:=0;
    exit;
  end;
	if (TheAffaire <> '' ) and (not AffaireExist (TheAffaire)) then
  begin
  	PgiBox ('Cette affaire n''existe pas');
    TForm(Ecran).ModalResult:=0;
  	result := false;
    exit;
  end;
end;

procedure EcritEnteteBord(X : TForm;TOBPiece : TOB);
var Fact : TFFacture;
		TOBE : TOB;
    QQ : TQuery;
    TheNum : integer;
begin
  TOBE := TOB.Create ('BDETETUDE',nil,-1);
  TRY
    Fact := TFFacture(X);
    TheNUM := 0;
    QQ := OpenSql('SELECT MAX(BDE_ORDRE) AS NUMMAX FROM BDETETUDE WHERE BDE_AFFAIRE="'+ TOBPiece.getValue('GP_AFFAIRE')+'" AND '+
                  'BDE_CLIENT="'+TOBPiece.getValue('GP_TIERS')+'" AND BDE_NATUREAUXI="CLI"',true);
    if not QQ.eof then TheNUm := QQ.findField('NUMMAX').asInteger;
    ferme (QQ);
    TOBE.putValue('BDE_NATUREAUXI','CLI');
    TOBE.putValue('BDE_AFFAIRE',TOBPiece.getValue('GP_AFFAIRE'));
    TOBE.putValue('BDE_CLIENT',TOBPiece.getValue('GP_TIERS'));
    TOBE.PutValue('BDE_QUALIFNAT','PRINC1');
    TOBE.putValue('BDE_ORDRE',TheNum+1);
    TOBE.putValue('BDE_NATUREDOC','001');
    TOBE.putValue('BDE_TYPE','002');
    TOBE.putValue('BDE_INDICE',0);
    TOBE.PutValue('BDE_NATUREPIECEG',TOBPiece.getValue('GP_NATUREPIECEG'));
    TOBE.PutValue('BDE_SOUCHE',TOBPiece.getValue('GP_SOUCHE'));
    TOBE.PutValue('BDE_NUMERO',TOBPiece.getValue('GP_NUMERO'));
    TOBE.PutValue('BDE_INDICEG',TOBPiece.getValue('GP_INDICEG'));
    TOBE.PutValue('BDE_SELECTIONNE','X');
    TOBE.PutValue('BDE_DESIGNATION',Fact.DESIGNATIONBord);
    TOBE.PutValue('BDE_DATEDEPART',Fact.DateDepartBord);
    TOBE.PutValue('BDE_DATEFIN',Fact.DateFinBord);
    TOBE.PutValue('BDE_PIECEASSOCIEE',EncodeRefPiece (TOBPiece));
    if not TOBE.InsertDB (nil,true) then
    begin
      MessageValid := 'Erreur mise à jour ENTETE BORDEREAU';
      V_PGI.IOError := OeUnknown;
    end;
  FINALLY
  	TOBE.free;
  END;
end;

Initialization
  registerclasses ( [ TOF_BTSELBORDEREAUX,TOF_BORDEREAU_AFF ] ) ;
  RegisterAglProc( 'DupliqueBordereaux', TRUE , 2, AglDupliqueBordereaux);
end.

