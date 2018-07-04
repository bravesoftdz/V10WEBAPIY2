unit UtofBTFAMRES_MUL;

interface
uses
  Classes,
{$IFDEF EAGLCLIENT}
	eMul,
  MainEAgl,
{$Else}
  mul,
  fe_Main,
  DB,
  DBGrids,
{$ENDIF}
  sysUtils,
  Ent1,
  HCtrls,
  UTOF,
  UTOM,
  HDB,
  Htb97;

Type
  TOF_BTFAMRES_MUL = Class (TOF)
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


procedure TOF_BTFAMRES_MUL.OnArgument (Argument : String );
var Critere : String;
    Valeur  : String;
    Champ   : String;
    X       : integer;
    CC : THValComboBox;
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

  Ecran.caption := 'Liste des types de ressources';
  UpdateCaption(Ecran);

  //Gestion Restriction Domaine et Etablissements
  CC:=THValComboBox(GetControl('HFR_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

  CC:=THValComboBox(GetControl('HFR_ETABLISSEMENT')) ;
  if CC<>Nil then PositionneEtabUser(CC) ;

end;

procedure TOF_BTFAMRES_MUL.InsertClick(Sender: TOBJect);
begin

  Agllancefiche ('BTP', 'BTFAMRES', '', '','ACTION=CREATION') ;
  TtoolBarButton97(GetControl('Bcherche')).Click;

end;

Procedure TOF_BTFAMRES_MUL.ControleChamp(Champ : String;Valeur : String);
Begin

end;


Procedure TOF_BTFAMRES_MUL.ControleCritere(Critere : String);
Begin

end;

procedure TOF_BTFAMRES_MUL.OnUpdate;
begin
inherited ;
{$IFNDEF EAGLCLIENT}
if not (Ecran is TFMul) then exit;
{$ENDIF EAGLCLIENT}
end;

procedure TOF_BTFAMRES_MUL.FlisteDblClick(Sender: TObject);
var Action 	: string;
    TypeRes	: String;
begin
{$IFDEF EAGLCLIENT}
  TFMul(ecran).Q.TQ.Seek(TFMul(ecran).FListe.Row-1) ;
  TypeRes:=TFMul(ecran).Q.FindField('HFR_FAMRES').AsString;
{$ELSE}
	TypeRes:=Fliste.datasource.dataset.FindField('HFR_FAMRES').AsString;
{$ENDIF}
Action:=GetControlText('XXAction');

if Action='ACTION=RECH' then
	begin
	TFMul(Ecran).Retour :=Typeres;
  TFMUL(Ecran).Close;
	end
else
	begin
   	AGLLanceFiche('BTP','BTFAMRES','',TypeRes,'');
    TtoolBarButton97(GetCOntrol('Bcherche')).Click;
	end;

end;

Initialization
registerclasses([TOF_BTFAMRES_MUL]);
end.
