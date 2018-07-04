{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 28/03/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTFACTTEXTES ()
Mots clefs ... : TOF;BTFACTTEXTES
*****************************************************************}
Unit BTFACTTEXTES;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
		 Fe_Main,
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
		 MaineAGL,
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
     HrichOle,
     AglInit,
     UTOF ;

Type
	TTypetexte = (TTtdebut,TTtFin);

  TOF_BTFACTTEXTES = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	Texte : THRichEditOLE;
    TOBPiece : TOB;
    TOBLIENOLE : TOB;
    procedure AfficheLeTexte;
  end ;

procedure ModifieTexteDebutFin ( TOBPiece,TOBLiensOle : TOB ; TypeTexte : TTypeTexte; ModeAction : TactionFiche);

Implementation
uses BTPUtil;

procedure ModifieTexteDebutFin ( TOBpiece,TOBLiensOle : TOB ; TypeTexte : TTypeTexte; ModeAction : TactionFiche);
var TOBParam : TOB;
		LEMODEACTION : string;
begin
	if ModeAction = taConsult then LEMODEACTION := 'ACTION=CONSULTATION'
  													else LEMODEACTION := 'ACTION=MODIFICATION';
	TOBParam := TOB.Create ('LE PARAM', nil,-1);
  if TypeTexte = tttDebut then TOBParam.AddChampSupValeur ('TYPETEXTE','1',false)
  												else TOBParam.AddChampSupValeur ('TYPETEXTE','2',false);
  TOBLiensOle.data := TOBPiece;
  TOBParam.Data := TOBLIensOle;
  TheTOB := TOBPAram;
  AglLancefiche ('BTP','BTFACTTEXTES','','',LEMODEACTION);
  TheTOB := nil;
  TOBLIensOle.data := nil;
  TOBParam.free;
end;



procedure TOF_BTFACTTEXTES.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTFACTTEXTES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTFACTTEXTES.OnUpdate ;
var T: TOB;
  NumPiece: string;
  Datepiece: TDatetime;
begin
  Inherited ;
  Numpiece := TOBPiece.GetValue('GP_NATUREPIECEG') + ':' +
              TOBPiece.GetValue('GP_SOUCHE') + ':' +
              IntToStr(TOBPiece.GetValue('GP_NUMERO')) + ':' +
              IntToStr(TOBPiece.GetValue('GP_INDICEG'));
  Datepiece := StrToDate(TOBPIECE.GetValue('GP_DATEPIECE'));
  T := TOBLIENOLE.FindFirst(['LO_TABLEBLOB', 'LO_IDENTIFIANT', 'LO_RANGBLOB'],
  													['GP', NumPiece, LaTOB.GetValue ('TYPETEXTE')], false);
  if T = nil then T := TOB.create('LIENSOLE', TOBLIENOLE, -1);
  if (Length(Texte.Text) <> 0) and (texte.Text <> #$D#$A) then
  begin
    T.PutValue('LO_TABLEBLOB', 'GP');
    T.PutValue('LO_QUALIFIANTBLOB', 'MEM');
    T.PutValue('LO_IDENTIFIANT', NumPiece);
    T.PutValue('LO_RANGBLOB', LaTOB.GetValue ('TYPETEXTE'));
    T.PutValue('LO_DATEBLOB', Datepiece);
    T.PutValue('LO_OBJET', ExRichtoString(Texte));
  end else
  begin
  	T.free;
  end;
end ;

procedure TOF_BTFACTTEXTES.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTFACTTEXTES.OnArgument (S : String ) ;
begin
  Inherited ;
  AppliqueFontDefaut (THRichEditOLE(ecran.FindComponent('TTEXTE')));
	if LaTOB.GetValue ('TYPETEXTE')= '1' then ecran.caption := TraduireMemoire('Texte de début')
  																		 else ecran.caption := TraduireMemoire('Texte de fin');
	Texte := THRichEditOle(GetControl('TTEXTE'));
  //
  TOBLIENOLE := TOB(LaTOB.data);
  TOBPiece := TOB(TOBLIENOLE.Data);
  AfficheLetexte;
end ;

procedure TOF_BTFACTTEXTES.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTFACTTEXTES.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTFACTTEXTES.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTFACTTEXTES.AfficheLeTexte;
var NumPiece : string;
		T : TOB;
begin
	texte.Clear;
  texte.Text := '';
  Numpiece := TOBPIece.getValue('GP_NATUREPIECEG') + ':' +
  						TOBPIece.getValue('GP_SOUCHE') + ':' +
              IntToStr(TOBPiece.getValue('GP_NUMERO')) + ':' +
              IntToStr(TOBPIECE.getValue('GP_INDICEG'));
  T := TOBLIENOLE.FindFirst(['LO_TABLEBLOB', 'LO_IDENTIFIANT', 'LO_RANGBLOB'], ['GP', NumPiece, LaTOB.GetValue ('TYPETEXTE')], false);
  if T = nil then
  begin
		T := TOB.create('LIENSOLE', TOBLIENOLE, -1);
    T.PutValue('LO_TABLEBLOB', 'GP');
    T.PutValue('LO_QUALIFIANTBLOB', 'MEM');
    T.PutValue('LO_IDENTIFIANT', NumPiece);
    T.PutValue('LO_RANGBLOB', LaTOB.GetValue ('TYPETEXTE'));
  end;
  StringToRich(TEXTE, T.GetValue('LO_OBJET'));
end;

Initialization
  registerclasses ( [ TOF_BTFACTTEXTES ] ) ;
end.

