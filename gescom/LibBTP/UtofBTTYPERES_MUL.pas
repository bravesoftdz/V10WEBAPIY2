unit UtofBTTYPERES_MUL;

interface
uses
  Classes,
{$IFDEF EAGLCLIENT}
	eMul,
  MaineAgl,
{$Else}
  mul,
  fe_Main,
  DB,
  DBGrids,
{$ENDIF}
  sysUtils,
  HCtrls,
  UTOF,
  UTOM,
  HDB,
  Htb97;

Type
  TOF_BTTYPERES_MUL = Class (TOF)
      public
      procedure OnArgument (Argument : String ) ; override ;
      procedure OnUpdate ; override ;

	Private

  	BInsert : TToolbarButton97;
		Fliste 	: THDbGrid;

  	procedure ControleChamp(Champ, Valeur: String);
  	procedure ControleCritere(Critere: String);

    //Procedure associée à une action sur un objet
    procedure FlisteDblClick (Sender : TObject);
    procedure InsertClick(Sender: TOBJect);

  end;

implementation


procedure TOF_BTTYPERES_MUL.OnArgument (Argument : String );
var Critere : String;
    Valeur  : String;
    Champ   : String;
    X       : integer;
begin
  Inherited ;

  //Récupération valeur de argument
	Critere:=(Trim(ReadTokenSt(Argument)));

  while (Critere <> '') do
    begin
      if Critere <> '' then
      begin
        X := pos (':', Critere) ;
        if x = 0 then
          X := pos ('=', Critere) ;
        if x <> 0 then
        begin
          Champ := copy (Critere, 1, X - 1) ;
          Valeur := Copy (Critere, X + 1, length (Critere) - X) ;
        	ControleChamp(champ, valeur);
				end
      end;
      ControleCritere(Critere);
      Critere := (Trim(ReadTokenSt(Argument)));
    end;

  BInsert := TToolbarButton97(ecran.FindComponent('BInsert'));
  BInsert.OnClick := InsertClick;

  Fliste := THDbGrid(GetControl('FLISTE'));
  Fliste.OnDblClick := FlisteDblClick;

  Ecran.caption := 'Liste des sous-types de ressources';
  UpdateCaption(Ecran);

end;

procedure TOF_BTTYPERES_MUL.InsertClick(Sender: TOBJect);
begin

  Agllancefiche ('BTP', 'BTTYPERES', '', '','ACTION=CREATION') ;
  TtoolBarButton97(GetControl('Bcherche')).Click;

end;

Procedure TOF_BTTYPERES_MUL.ControleChamp(Champ : String;Valeur : String);
Begin

end;


Procedure TOF_BTTYPERES_MUL.ControleCritere(Critere : String);
Begin

end;

procedure TOF_BTTYPERES_MUL.OnUpdate;
{$IFNDEF EAGLCLIENT}
var i:integer;
    F : TFMUL;
    h:string;
    s:string;
{$ENDIF EAGLCLIENT}
begin
inherited ;
{$IFNDEF EAGLCLIENT}
if not (Ecran is TFMul) then exit;
{$ENDIF EAGLCLIENT}
end;

procedure TOF_BTTYPERES_MUL.FlisteDblClick(Sender: TObject);
var Action 	: string;
    TypeRes	: String;
    Famres  : String;
begin

TypeRes:=Fliste.datasource.dataset.FindField('BTR_TYPRES').AsString;
FamRes:=Fliste.datasource.dataset.FindField('BTR_FAMRES').AsString;

Action:=GetControlText('XXAction');

if Action='ACTION=RECH' then
	begin
	TFMul(Ecran).Retour :=Typeres;
  TFMUL(Ecran).Close;
	end
else
	begin
   	AGLLanceFiche('BTP','BTTYPERES','','','TYPRES=' + TypeRes + ';FAMRES=' + Famres + '; ACTION=MODIFICATION');
    TtoolBarButton97(GetControl('Bcherche')).Click;
	end;

end;

Initialization
registerclasses([TOF_BTTYPERES_MUL]);
end.
