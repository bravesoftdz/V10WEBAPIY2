unit UFactgestionAff;

interface

uses  UTOB,
			Hent1,
      Ent1,
      HCtrls,
      forms,
{$IFDEF EAGLCLIENT}
     UtileAGL,MaineAGL,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}EdtREtat,FE_Main,uPDFBatch,
{$ENDIF}
			BtpUtil,
      SysUtils,
      graphics,
      classes,
      controls,
      UlibWindows,
      UtilFichiers,
      Windows, Messages, StdCtrls;

type

  TInfoAffiche = class (Tobject)
    fBrushCol : TColor;
    fFont      : TFont;
  public
    constructor create (XX:TForm);
    destructor destroy; override;
    property brush : Tcolor read fBrushCol write fBrushCol ;
    property Font : TFont read fFont write fFont ;
  end;

  TTypeAffichage = (TtaDefaut,TtaNormal);

	TAffichageDoc = class (Tobject)
  	private
      fSystemFont : Tfont;
			TOBPAram : TOB;
      fTypeaffichage : TTypeAffichage;
      Fpar1      : TinfoAffiche;
      Fpar2      : TinfoAffiche;
      Fpar3      : TinfoAffiche;
      Fpar4      : TinfoAffiche;
      Fpar5      : TinfoAffiche;
      Fpar6      : TinfoAffiche;
      Fpar7      : TinfoAffiche;
      Fpar8      : TinfoAffiche;
      Fpar9      : TinfoAffiche;
      fparOuv    : TinfoAffiche;
      fParSDet   : TinfoAffiche;
      fParVar    : TInfoAffiche;
      fParStot   : TInfoAffiche;
      fparLig    : TInfoAffiche;
      fparCom    : TInfoAffiche;
      fDecalPar  : integer;
      fDecalSDet : integer;
      Faction    : TActionFiche;
    	procedure SetTypeaffichage(const Value: TTypeAffichage);
      procedure GetValeurs ;
    public
      constructor create (FF: TForm);
      destructor destroy; override;
      property Paragraphe1 : TinfoAffiche read fPar1 write Fpar1;
      property Paragraphe2 : TinfoAffiche read fPar2 write Fpar2;
      property Paragraphe3 : TinfoAffiche read fPar3 write Fpar3;
      property Paragraphe4 : TinfoAffiche read fPar3 write Fpar4;
      property Paragraphe5 : TinfoAffiche read fPar5 write Fpar5;
      property Paragraphe6 : TinfoAffiche read fPar6 write Fpar6;
      property Paragraphe7 : TinfoAffiche read fPar7 write Fpar7;
      property Paragraphe8 : TinfoAffiche read fPar8 write Fpar8;
      property Paragraphe9 : TinfoAffiche read fPar9 write Fpar9;
      property Ouvrage     : TInfoAffiche read fparOuv write fparOuv;
      property SousDetail  : TInfoAffiche read fParSDet write FparSdet;
      property Variante    : TInfoAffiche read fParVar write FparVar;
      property SousTotaux  : TInfoAffiche read fParStot write FparStot;
      property Ligne       : TInfoAffiche read fParlig write FparLig;
      property Commentaire : TInfoAffiche read fParCom write FparCom;
      property DecalParag  : integer read fDecalPar write FDecalPar;
      property DecalSousDetail : integer read fDecalSDet write fDecalSDet;
      property gestion : TTypeAffichage read fTypeaffichage write SetTypeaffichage;
      property Action : TActionFiche read Faction write Faction;
      procedure SetValeurs;
      procedure Enregistrevaleurs;
  end;

procedure DupliqueValeurs (Source,Destination : TAffichageDoc);
procedure AppliqueStyleArticle (Canvas : Tcanvas; Current : TAffichageDoc; ChangeColor : Boolean=true);
procedure AppliqueStyleCommentaire (Canvas : Tcanvas; Current : TAffichageDoc; ChangeColor : Boolean=true);
procedure AppliqueStyleDetail (Canvas : Tcanvas; Current : TAffichageDoc; ChangeColor : Boolean=true);
procedure AppliqueStyleParag (canvas : Tcanvas; Current : TAffichageDoc; Ordre : integer; ChangeColor : Boolean=true);
procedure AppliqueStyleSousTot (Canvas : Tcanvas; Current : TAffichageDoc; ChangeColor : Boolean=true);
procedure AppliqueStyleVariante (Canvas : Tcanvas; Current : TAffichageDoc; ChangeColor : Boolean=true);
procedure AppliqueStyleOuvrage (Canvas : Tcanvas; Current : TAffichageDoc; ChangeColor : Boolean=true);
function GetDecalageSousDetail (Current : TaffichageDoc) : integer;
function GetDecalageParagraphe (Current : TaffichageDoc) : integer;

implementation


procedure AppliqueStyleCommentaire (Canvas : Tcanvas; Current : TAffichageDoc; ChangeColor : Boolean=true);
begin
//  Canvas.Brush.Color := Current.Commentaire.brush;
  canvas.Font.Name := Current.Commentaire.Font.Name ;
  canvas.Font.size  := Current.Commentaire.Font.Size ;
  if ChangeColor then canvas.Font.Color := Current.Commentaire.Font.Color ;
  canvas.Font.Style := Current.Commentaire.Font.Style ;
end;

procedure AppliqueStyleArticle (Canvas : Tcanvas; Current : TAffichageDoc; ChangeColor : Boolean=true);
begin
//  Canvas.Brush.Color := Current.Ligne.brush;
  canvas.Font.Name := Current.Ligne.Font.Name ;
  canvas.Font.Size := Current.Ligne.Font.Size ;
  if ChangeColor then canvas.Font.Color := Current.Ligne.Font.Color ;
  canvas.Font.Style := Current.Ligne.Font.Style ;
end;

procedure AppliqueStyleOuvrage (Canvas : Tcanvas; Current : TAffichageDoc; ChangeColor : Boolean=true);
begin
//  Canvas.Brush.Color := Current.Ouvrage.brush;
  canvas.Font.Name := Current.Ouvrage.Font.Name ;
  canvas.Font.Size := Current.Ouvrage.Font.Size ;
  if ChangeColor then canvas.Font.Color := Current.Ouvrage.Font.Color ;
  canvas.Font.Style := Current.Ouvrage.Font.Style ;
end;

procedure AppliqueStyleVariante (Canvas : Tcanvas; Current : TAffichageDoc; ChangeColor : Boolean=true);
begin
//  Canvas.Brush.Color := Current.Variante.brush;
  canvas.Font.Name := Current.Variante.Font.Name ;
  canvas.Font.Size := Current.Variante.Font.Size ;
  if ChangeColor then canvas.Font.Color := Current.Variante.Font.Color ;
  canvas.Font.Style := Current.Variante.Font.Style ;
end;

procedure AppliqueStyleDetail (Canvas : Tcanvas; Current : TAffichageDoc; ChangeColor : Boolean=true);
begin
//  Canvas.Brush.Color := Current.SousDetail.brush;
  canvas.Font.Name := Current.SousDetail.Font.Name ;
  canvas.Font.Size := Current.SousDetail.Font.Size ;
  if ChangeColor then canvas.Font.Color := Current.SousDetail.Font.Color ;
  canvas.Font.Style := Current.SousDetail.Font.Style ;
end;

procedure AppliqueStyleSousTot (Canvas : Tcanvas; Current : TAffichageDoc; ChangeColor : Boolean=true);
begin
//  Canvas.Brush.Color := Current.SousTotaux.brush;
  canvas.Font.Name := Current.SousTotaux.Font.Name ;
  canvas.Font.Size := Current.SousTotaux.Font.Size ;
  if ChangeColor then canvas.Font.Color := Current.SousTotaux.Font.Color ;
  canvas.Font.Style := Current.SousTotaux.Font.Style ;
end;

procedure AppliqueStyleParag (canvas : Tcanvas; Current : TAffichageDoc; Ordre : integer; ChangeColor : Boolean=true);
begin
	if Ordre = 1 then
  begin
//    Canvas.Brush.Color := Current.Paragraphe1.brush;
    canvas.Font.Name := Current.Paragraphe1.Font.Name ;
    canvas.Font.size  := Current.Paragraphe1.Font.Size ;
    if ChangeColor then canvas.Font.Color := Current.Paragraphe1.Font.Color ;
    canvas.Font.Style := Current.Paragraphe1.Font.Style ;
  end else if ordre = 2 then
  begin
//    Canvas.Brush.Color := Current.Paragraphe2.brush;
    canvas.Font.Name := Current.Paragraphe2.Font.Name ;
    canvas.Font.size  := Current.Paragraphe2.Font.Size ;
    if ChangeColor then canvas.Font.Color := Current.Paragraphe2.Font.Color ;
    canvas.Font.Style := Current.Paragraphe2.Font.Style ;
  end else if ordre = 3 then
	begin
//    Canvas.Brush.Color := Current.Paragraphe3.brush;
    canvas.Font.Name := Current.Paragraphe3.Font.Name ;
    canvas.Font.size  := Current.Paragraphe3.Font.Size ;
    if ChangeColor then canvas.Font.Color := Current.Paragraphe3.Font.Color ;
    canvas.Font.Style := Current.Paragraphe3.Font.Style ;
  end else if ordre = 4 then
  begin
//    Canvas.Brush.Color := Current.Paragraphe4.brush;
    canvas.Font.Name := Current.Paragraphe4.Font.Name ;
    canvas.Font.size  := Current.Paragraphe4.Font.Size ;
    if ChangeColor then canvas.Font.Color := Current.Paragraphe4.Font.Color ;
    canvas.Font.Style := Current.Paragraphe4.Font.Style ;
  end else if ordre = 5 then
  begin
//    Canvas.Brush.Color := Current.Paragraphe5.brush;
    canvas.Font.Name := Current.Paragraphe5.Font.Name ;
    canvas.Font.size  := Current.Paragraphe5.Font.Size ;
    if ChangeColor then canvas.Font.Color := Current.Paragraphe5.Font.Color ;
    canvas.Font.Style := Current.Paragraphe5.Font.Style ;
  end else if ordre = 6 then
  begin
//    Canvas.Brush.Color := Current.Paragraphe6.brush;
    canvas.Font.Name := Current.Paragraphe6.Font.Name ;
    canvas.Font.size  := Current.Paragraphe6.Font.Size ;
    if ChangeColor then canvas.Font.Color := Current.Paragraphe6.Font.Color ;
    canvas.Font.Style := Current.Paragraphe6.Font.Style ;
  end else if ordre = 7 then
  begin
//    Canvas.Brush.Color := Current.Paragraphe7.brush;
    canvas.Font.Name := Current.Paragraphe7.Font.Name ;
    canvas.Font.size := Current.Paragraphe7.Font.size ;
    if ChangeColor then canvas.Font.Color := Current.Paragraphe7.Font.Color ;
    canvas.Font.Style := Current.Paragraphe7.Font.Style ;
  end else if ordre = 8 then
  begin
//    Canvas.Brush.Color := Current.Paragraphe8.brush;
    canvas.Font.Name := Current.Paragraphe8.Font.Name ;
    canvas.Font.size := Current.Paragraphe8.Font.size ;
    if ChangeColor then canvas.Font.Color := Current.Paragraphe8.Font.Color ;
    canvas.Font.Style := Current.Paragraphe8.Font.Style ;
  end else if ordre = 9 then
  begin
//    Canvas.Brush.Color := Current.Paragraphe9.brush;
    canvas.Font.Name := Current.Paragraphe9.Font.Name ;
    canvas.Font.size := Current.Paragraphe9.Font.size ;
    if ChangeColor then canvas.Font.Color := Current.Paragraphe9.Font.Color ;
    canvas.Font.Style := Current.Paragraphe9.Font.Style ;
  end;
end;

function GetDecalageSousDetail (Current : TaffichageDoc) : integer;
begin
	result := Current.DecalSousDetail;
end;

function GetDecalageParagraphe (Current : TaffichageDoc) : integer;
begin
	result := Current.DecalParag;
end;

procedure DupliqueValeurs (Source,Destination : TAffichageDoc);
begin
  Destination.Paragraphe1.brush := Source.Paragraphe1.brush;
  Destination.Paragraphe1.Font.color  := Source.Paragraphe1.Font.Color;
  Destination.Paragraphe1.Font.Name   := Source.Paragraphe1.Font.name;
  Destination.Paragraphe1.Font.Style   := Source.Paragraphe1.Font.Style;
  //
  Destination.Paragraphe2.brush := Source.Paragraphe2.brush;
  Destination.Paragraphe2.Font.Color  := Source.Paragraphe2.Font.Color;
  Destination.Paragraphe2.Font.Name   := Source.Paragraphe2.Font.name;
  Destination.Paragraphe2.Font.Style   := Source.Paragraphe2.Font.Style;
  //
  Destination.Paragraphe3.brush := Source.Paragraphe3.brush;
  Destination.Paragraphe3.Font.Color  := Source.Paragraphe3.Font.Color;
  Destination.Paragraphe3.Font.Name   := Source.Paragraphe3.Font.name;
  Destination.Paragraphe3.Font.Style   := Source.Paragraphe3.Font.Style;
  //
  Destination.Paragraphe4.brush := Source.Paragraphe4.brush;
  Destination.Paragraphe4.Font.Color  := Source.Paragraphe4.Font.Color;
  Destination.Paragraphe4.Font.Name   := Source.Paragraphe4.Font.name;
  Destination.Paragraphe4.Font.Style   := Source.Paragraphe4.Font.Style;
  //
  Destination.Paragraphe5.brush := Source.Paragraphe5.brush;
  Destination.Paragraphe5.Font.Color  := Source.Paragraphe5.Font.Color;
  Destination.Paragraphe5.Font.Name   := Source.Paragraphe5.Font.name;
  Destination.Paragraphe5.Font.Style   := Source.Paragraphe5.Font.Style;
  //
  Destination.Paragraphe6.brush := Source.Paragraphe6.brush;
  Destination.Paragraphe6.Font.Color  := Source.Paragraphe6.Font.Color;
  Destination.Paragraphe6.Font.Name   := Source.Paragraphe6.Font.name;
  Destination.Paragraphe6.Font.Style   := Source.Paragraphe6.Font.Style;
  //
  Destination.Paragraphe7.brush := Source.Paragraphe7.brush;
  Destination.Paragraphe7.Font.Color  := Source.Paragraphe7.Font.Color;
  Destination.Paragraphe7.Font.Name   := Source.Paragraphe7.Font.name;
  Destination.Paragraphe7.Font.Style   := Source.Paragraphe7.Font.Style;
  //
  Destination.Paragraphe8.brush := Source.Paragraphe8.brush;
  Destination.Paragraphe8.Font.Color  := Source.Paragraphe8.Font.Color;
  Destination.Paragraphe8.Font.Name   := Source.Paragraphe8.Font.name;
  Destination.Paragraphe8.Font.Style   := Source.Paragraphe8.Font.Style;
  //
  Destination.Paragraphe9.brush := Source.Paragraphe9.brush;
  Destination.Paragraphe9.Font.Color  := Source.Paragraphe9.Font.Color;
  Destination.Paragraphe9.Font.Name   := Source.Paragraphe9.Font.name;
  Destination.Paragraphe9.Font.Style   := Source.Paragraphe9.Font.Style;
  //
  Destination.Ouvrage.brush := Source.Ouvrage.brush;
  Destination.Ouvrage.Font.Color  := Source.Ouvrage.Font.Color;
  Destination.Ouvrage.Font.Name   := Source.Ouvrage.Font.name;
  Destination.Ouvrage.Font.Style   := Source.Ouvrage.Font.Style;
  //
  Destination.SousDetail.brush := Source.SousDetail.brush;
  Destination.SousDetail.Font.Color  := Source.SousDetail.Font.Color;
  Destination.SousDetail.Font.Name   := Source.SousDetail.Font.name;
  Destination.SousDetail.Font.Style   := Source.SousDetail.Font.Style;
  //
  Destination.Variante.brush := Source.Variante.brush;
  Destination.Variante.Font.Color  := Source.Variante.Font.Color;
  Destination.Variante.Font.Name   := Source.Variante.Font.name;
  Destination.Variante.Font.Style   := Source.Variante.Font.Style;
  //
  Destination.SousTotaux.brush := Source.SousTotaux.brush;
  Destination.SousTotaux.Font.Color  := Source.SousTotaux.Font.Color;
  Destination.SousTotaux.Font.Name   := Source.SousTotaux.Font.name;
  Destination.SousTotaux.Font.Style   := Source.SousTotaux.Font.Style;
  //
  Destination.Ligne.brush := Source.Ligne.brush;
  Destination.Ligne.Font.Color  := Source.Ligne.Font.Color;
  Destination.Ligne.Font.Name   := Source.Ligne.Font.name;
  Destination.Ligne.Font.Style   := Source.Ligne.Font.Style;
  //
  Destination.Commentaire.brush := Source.Commentaire.brush;
  Destination.Commentaire.Font.Color  := Source.Commentaire.Font.Color;
  Destination.Commentaire.Font.Name   := Source.Commentaire.Font.name;
  Destination.Commentaire.Font.Style   := Source.Commentaire.Font.Style;
  //
  Destination.DecalParag := Source.DecalParag;
  Destination.DecalSousDetail := Source.DecalSousDetail;
end;

{ TAffichageDoc }

procedure TAffichageDoc.GetValeurs;
var QQ : TQuery;
		fgestion : string;
begin
  if fTypeaffichage = TtaDefaut then fgestion := 'DEFAUT'
  															else fgestion := 'ACTIF';
  QQ := OpenSql ('SELECT * FROM BPARAMDOC WHERE BD2_MODE="'+fgestion+'"',true,1,'',true);
  if not QQ.eof then
  begin
  	TOBPAram := TOB.Create ('BPARAMDOC',nil,-1);
    TOBPAram.SelectDB ('',QQ);
  end;
  ferme (QQ);
  //
	fPar1.fBrushCol := TOBParam.GetInteger('BD2_BRUSHPAR1');
  DecodeFont(TOBParam.GetString('BD2_FONTPAR1'),fPar1.fFont);
	fPar1.Font.Color  := TOBParam.GetInteger('BD2_COLORPAR1');
  //
	fPar2.fBrushCol := TOBParam.GetInteger('BD2_BRUSHPAR2');
  DecodeFont(TOBParam.GetString('BD2_FONTPAR2'),fPar2.fFont);
	fPar2.Font.Color  := TOBParam.GetInteger('BD2_COLORPAR2');
  //
	fPar3.fBrushCol := TOBParam.GetInteger('BD2_BRUSHPAR3');
  DecodeFont(TOBParam.GetString('BD2_FONTPAR3'),fPar3.fFont);
	fPar3.Font.Color  := TOBParam.GetInteger('BD2_COLORPAR3');
  //
	fPar4.fBrushCol := TOBParam.GetInteger('BD2_BRUSHPAR4');
  DecodeFont(TOBParam.GetString('BD2_FONTPAR4'),fPar4.fFont);
	fPar4.Font.Color  := TOBParam.GetInteger('BD2_COLORPAR4');
  //
	fPar5.fBrushCol := TOBParam.GetInteger('BD2_BRUSHPAR5');
  DecodeFont(TOBParam.GetString('BD2_FONTPAR5'),fPar5.fFont);
	fPar5.Font.Color  := TOBParam.GetInteger('BD2_COLORPAR5');
  //
	fPar6.fBrushCol := TOBParam.GetInteger('BD2_BRUSHPAR6');
  DecodeFont(TOBParam.GetString('BD2_FONTPAR6'),fPar6.fFont);
	fPar6.FOnt.Color  := TOBParam.GetInteger('BD2_COLORPAR6');
  //
	fPar7.fBrushCol := TOBParam.GetInteger('BD2_BRUSHPAR7');
  DecodeFont(TOBParam.GetString('BD2_FONTPAR7'),fPar7.fFont);
	fPar7.Font.Color  := TOBParam.GetInteger('BD2_COLORPAR7');
  //
	fPar8.fBrushCol := TOBParam.GetInteger('BD2_BRUSHPAR8');
  DecodeFont(TOBParam.GetString('BD2_FONTPAR8'),fPar8.fFont);
	fPar8.Font.Color  := TOBParam.GetInteger('BD2_COLORPAR8');
  //
	fPar9.fBrushCol := TOBParam.GetInteger('BD2_BRUSHPAR9');
  DecodeFont(TOBParam.GetString('BD2_FONTPAR9'),fPar9.fFont);
	fPar9.Font.Color  := TOBParam.GetInteger('BD2_COLORPAR9');
  //
	fparOuv.fBrushCol := TOBParam.GetInteger('BD2_BRUSHOUV');
  DecodeFont(TOBParam.GetString('BD2_FONTOUV'),fParOuv.fFont);
	fParOuv.Font.Color  := TOBParam.GetInteger('BD2_COLOROUV');
  //
	fParSDet.fBrushCol := TOBParam.GetInteger('BD2_BRUSHDETAIL');
  DecodeFont(TOBParam.GetString('BD2_FONTDETAIL'),fParSDet.fFont);
	fParSDet.Font.Color  := TOBParam.GetInteger('BD2_COLORDETAIL');
  //
	fParStot.fBrushCol := TOBParam.GetInteger('BD2_BRUSHSTOT');
  DecodeFont(TOBParam.GetString('BD2_FONTSTOT'),fParStot.fFont);
	fParStot.Font.Color  := TOBParam.GetInteger('BD2_COLORSTOT');
  //
	fParVar.fBrushCol := TOBParam.GetInteger('BD2_BRUSHVAR');
  DecodeFont(TOBParam.GetString('BD2_FONTVAR'),fParVar.fFont);
	fParVar.Font.Color  := TOBParam.GetInteger('BD2_COLORVAR');
  //
	fParLig.fBrushCol := TOBParam.GetInteger('BD2_BRUSHLIG');
  DecodeFont(TOBParam.GetString('BD2_FONTLIG'),fParLig.fFont);
	fParLig.Font.Color  := TOBParam.GetInteger('BD2_COLORLIG');
  //
	fParCom.fBrushCol := TOBParam.GetInteger('BD2_BRUSHCOM');
  DecodeFont(TOBParam.GetString('BD2_FONTCOM'),fParCom.fFont);
	fParCom.Font.Color  := TOBParam.GetInteger('BD2_COLORCOM');
	//
  fDecalPar := TOBParam.GetInteger('BD2_DECALPARAG');
  fDecalSDet := TOBParam.GetInteger('BD2_DECALSDETAIL');
  //
end;

constructor TAffichageDoc.create (FF : Tform);
begin
  fSystemFont := Tfont.Create;
  TOBParam := TOB.create ('BPARAMDOC',nil,-1);
  fPar1 := TInfoAffiche.create(FF);
  fPar2 := TInfoAffiche.create(FF);
  fPar3 := TInfoAffiche.create(FF);
  fPar4 := TInfoAffiche.create(FF);
  fPar5 := TInfoAffiche.create(FF);
  fPar6 := TInfoAffiche.create(FF);
  fPar7 := TInfoAffiche.create(FF);
  fPar8 := TInfoAffiche.create(FF);
  fPar9 := TInfoAffiche.create(FF);
  fparOuv := TInfoAffiche.create(FF);
  fParSDet := TInfoAffiche.create(FF);
  fParStot := TInfoAffiche.create(FF);
  fParVar  := TInfoAffiche.create(FF);
  fParlig  := TInfoAffiche.create(FF);
  fParCom  := TInfoAffiche.create(FF);
end;

destructor TAffichageDoc.destroy;
begin
  fPar1.free;
  fPar2.free;
  fPar3.free;
  fPar4.free;
  fPar5.free;
  fPar6.free;
  fPar7.free;
  fPar8.free;
  fPar9.free;
  fparOuv.free;
  fParSDet.free;
  fParStot.free;
  fParVar.free;
  fParLig.free;
  fParCom.free;
  TOBParam.free;
  inherited;
end;

procedure TAffichageDoc.SetTypeaffichage(const Value: TTypeAffichage);
begin
  fTypeaffichage := Value;
	GetValeurs;
end;

procedure TAffichageDoc.SetValeurs;
begin
	TOBParam.SetInteger('BD2_BRUSHPAR1',fPar1.fBrushCol);
  TOBParam.SetString('BD2_FONTPAR1',EncodeFont (fPar1.fFont));
	TOBParam.SetInteger('BD2_COLORPAR1',fPar1.Font.Color);
  //
	TOBParam.SetInteger('BD2_BRUSHPAR2',fPar2.fBrushCol);
  TOBParam.SetString('BD2_FONTPAR2',EncodeFont (fPar2.fFont));
	TOBParam.SetInteger('BD2_COLORPAR2',fPar2.Font.Color);
  //
	TOBParam.SetInteger('BD2_BRUSHPAR3',fPar3.fBrushCol);
  TOBParam.SetString('BD2_FONTPAR3',EncodeFont (fPar3.fFont));
	TOBParam.SetInteger('BD2_COLORPAR3',fPar3.Font.Color);
  //
	TOBParam.SetInteger('BD2_BRUSHPAR4',fPar4.fBrushCol);
  TOBParam.SetString('BD2_FONTPAR4',EncodeFont(fPar4.fFont));
	TOBParam.SetInteger('BD2_COLORPAR4',fPar4.Font.Color);
  //
	TOBParam.SetInteger('BD2_BRUSHPAR5',fPar5.fBrushCol);
  TOBParam.SetString('BD2_FONTPAR5',EncodeFont(fPar5.fFont));
	TOBParam.SetInteger('BD2_COLORPAR5',fPar5.Font.Color);
  //
	TOBParam.SetInteger('BD2_BRUSHPAR6',fPar6.fBrushCol);
  TOBParam.SetString('BD2_FONTPAR6',EncodeFont(fPar6.fFont));
	TOBParam.SetInteger('BD2_COLORPAR6',fPar6.Font.Color);
  //
	TOBParam.SetInteger('BD2_BRUSHPAR7',fPar7.fBrushCol);
  TOBParam.SetString('BD2_FONTPAR7',EncodeFont(fPar7.fFont));
	TOBParam.SetInteger('BD2_COLORPAR7',fPar7.Font.Color);
  //
	TOBParam.SetInteger('BD2_BRUSHPAR8',fPar8.fBrushCol);
  TOBParam.SetString('BD2_FONTPAR8',EncodeFont(fPar8.fFont));
	TOBParam.SetInteger('BD2_COLORPAR8',fPar8.Font.Color);
  //
	TOBParam.SetInteger('BD2_BRUSHPAR9',fPar9.fBrushCol);
  TOBParam.SetString('BD2_FONTPAR9',EncodeFont(fPar9.fFont));
	TOBParam.SetInteger('BD2_COLORPAR9',fPar9.Font.Color);
  //
  TOBParam.SetInteger('BD2_BRUSHOUV',fparOuv.fBrushCol);
  TOBParam.SetString('BD2_FONTOUV',EncodeFont(fParOuv.fFont));
	TOBParam.SetInteger('BD2_COLOROUV',fParOuv.Font.Color);
  //
	TOBParam.SetInteger('BD2_BRUSHDETAIL',fParSDet.fBrushCol);
  TOBParam.SetString('BD2_FONTDETAIL',EncodeFont(fParSDet.fFont));
  TOBParam.SetInteger('BD2_COLORDETAIL',fParSDet.Font.Color );
  //
	TOBParam.SetInteger('BD2_BRUSHSTOT',fParStot.fBrushCol);
  TOBParam.SetString('BD2_FONTSTOT',EncodeFont(fParStot.fFont));
	TOBParam.SetInteger('BD2_COLORSTOT',fParStot.Font.Color);
  //
	TOBParam.SetInteger('BD2_BRUSHVAR',fParVar.fBrushCol);
  TOBParam.SetString('BD2_FONTVAR',EncodeFont(fParVar.fFont));
	TOBParam.SetInteger('BD2_COLORVAR',fParVar.Font.Color);
  //
	TOBParam.SetInteger('BD2_BRUSHLIG',fparLig.fBrushCol);
  TOBParam.SetString('BD2_FONTLIG',EncodeFont(fparLig.fFont));
	TOBParam.SetInteger('BD2_COLORLIG',fparLig.Font.Color);
  //
	TOBParam.SetInteger('BD2_BRUSHCOM',fparCom.fBrushCol);
  TOBParam.SetString('BD2_FONTCOM',EncodeFont(fparCom.fFont));
	TOBParam.SetInteger('BD2_COLORCOM',fparCom.Font.Color);
	//
  TOBParam.SetInteger('BD2_DECALPARAG',fDecalPar );
  TOBParam.SetInteger('BD2_DECALSDETAIL',fDecalSDet);
  //

end;

procedure TAffichageDoc.Enregistrevaleurs;
begin
  SetValeurs;
  TOBParam.SetAllModifie(true);
	TOBPAram.InsertOrUpdateDB(false);
end;

{ TInfoAffiche }

constructor TInfoAffiche.create (XX : TForm);
begin
  fFont := Tfont.Create;
  fFont.Name := XX.Font.name;
  fFont.Style := XX.Font.Style;
  fFont.Size := XX.Font.Size;
end;

destructor TInfoAffiche.destroy;
begin
  fFont.free;
  inherited;
end;

end.
