unit TofSuiviClient;

(*
procedure TForm1.TobViewer1DblClick(Sender: TObject);
begin
    with TTobViewer(Sender) do
        begin
        // CurrentRow représente le n° de ligne sur laquelle s'est produit le Double-Click
        // La première ligne est toujours 0. Les numéros suivent l'ordre d'affichage de haut en bas
        // (lié aux tris) et ne tiennent pas compte des lignes ajoutées par un groupement (Header/Footer).
        ShowMessage(IntToStr( CurrentRow ));

        // CurrentCol représente le n° de la colonne sur laquelle s'est produit le Double-Click.
        // La première colone créée porte le n° 0.  Lorsque les colonnes sont permutées par
        // "Glisser/Lâcher", elles conservent leurs n°.
        ShowMessage(IntToStr( CurrentCol ));
 
        // --- Récupération d'un nom de colonne à partir de son index (numéro)
        // En règle générale, le nom d'une colonne correspond au nom du champ associé.
        // Pour Afficher le nom de la colonne:
        ShowMessage( ColName[CurrentCol] );
 
        // --- Récupération d'un index de colonne à partir de son nom
        ShowMessage(IntToStr( ColIndex["Nom_de_mon_champ"] ));
 
        // --- Récupération de la valeur d'une cellule TEXTE
        ShowMessage( AsString[CurrentCol, CurrentRow] );
        ShowMessage( AsString[ColIndex["Nom_de_mon_champ"], CurrentRow] );
 
        // --- Récupération de la valeur d'une cellule BOOLENNE
        if AsBoolean[ColIndex["Nom_de_mon_champ"], CurrentRow] then
            ShowMessage( 'Vrai' )
        else
            ShowMessage('Faux');

        // --- Récupération de la valeur d'une cellule ENTIER
        ShowMessage(IntToStr( AsInteger[ColIndex["Nom_de_mon_champ"], CurrentRow] ));

        // --- Récupération de la valeur d'une cellule RELLE
        ShowMessage(FloatToStr( AsDouble[ColIndex["Nom_de_mon_champ"], CurrentRow] ));

        // --- Récupération de la valeur d'une cellule DATE
        ShowMessage(DatetimeToStr( AsDate[ColIndex["Nom_de_mon_champ"], CurrentRow] ));

        end;
end;
*)

interface

uses Classes, StdCtrls, SysUtils,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     comctrls, forms, ParamSoc,		// GetParamSocSecur YMO
     UTof, HCtrls, QRS1, Ent1,  TofMeth, HTB97, UTobView ;

type
  TOF_SUIVICLIENT = class(TOF_Meth)
  private
    gen1, gen2, aux1, aux2, Jal1, Jal2: THEdit;
    paie, Exo, Etab, Devise, sens : THValComboBox;
    Ech1, Ech2 : THEdit;
    xxWhere: THEdit;
    pages: TPageControl;
    DateD, DateF: TDatetime;
    Tidc, Rib: TCheckBox ;
    PzLibre: TTabSheet ;
    TViewer: TTobViewer;

    procedure CompteOnExit(Sender: TObject) ;
    procedure JalOnExit(Sender: TObject) ;
    procedure ExoOnChange(Sender: TObject) ;
    procedure DateOnExit(Sender: TObject) ;
    procedure SensOnChange(Sender: TObject) ;
    procedure TidcOnClick(Sender: TObject) ;
    procedure AuxiElipsisClick(Sender : TObject);
    procedure InitTous;
  public
    procedure OnArgument(Arguments : string) ; override ;
    procedure OnUpdate ; override ;
    procedure OnNew ; override ;
    procedure OnLoad ; override ;
  end ;

//procedure AGLSuiviClient(SCTob: Tob) ;
procedure AGLSuiviClient ;

implementation

uses Dialogs, HEnt1, UTob, Fe_main, Stat, UTofMulParamGen; {26/04/07 YMO F5 sur Auxiliaire }

var TheData : TObject ;

procedure AGLSuiviClient ; //(SCTob: Tob) ;
var SCTob: Tob ; Q: TQuery ;
begin
(*
Q:=OpenSQL('SELECT E_EXERCICE, E_GENERAL, E_AUXILIAIRE, E_LIBELLE, E_JOURNAL,'
          +'E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_DEBIT, E_CREDIT,'
          +'E_DEBITEURO, E_CREDITEURO, E_DEBITDEV, E_CREDITDEV, E_REFINTERNE,'
          +'E_QUALIFPIECE, E_VALIDE, E_DATEREFEXTERNE, E_DATECREATION,'
          +'E_DATEMODIF, E_REFLIBRE, E_CONTREPARTIEGEN,'
          +'E_CONTREPARTIEAUX, E_COUVERTURE, E_LETTRAGE, E_MODEPAIE,'
          +'E_DATEECHEANCE, E_UTILISATEUR, E_TIERSPAYEUR, E_ECRANOUVEAU,'
          +'E_COUVERTUREDEV, E_NUMECHE, E_SAISIEEURO, E_DEVISE, E_RIB,'
          +'E_ETATLETTRAGE, E_ANA, E_BANQUEPREVI, E_TAUXDEV,'
          +'E_REFRELEVE,E_NOMLOT, E_COUVERTUREEURO, E_COTATION,'
          +'E_CODEACCEPT, E_PERIODE, E_NUMTRAITECHQ, E_NUMENCADECA,'
          +'G_LIBELLE, G_NATUREGENE, T_NATUREAUXI, T_LIBELLE, T_TABLE0,'
          +'T_TABLE1, T_TABLE2, T_TABLE3, T_TABLE4, T_TABLE5, T_TABLE6, T_TABLE7,'
          +'T_TABLE8, T_TABLE9 FROM SUIVIMP', TRUE) ;
*)
Q:=OpenSQL('SELECT E_EXERCICE, E_GENERAL, E_AUXILIAIRE, E_LIBELLE, E_JOURNAL,'
          +'E_DATECOMPTABLE FROM ECRITURE', TRUE) ;
SCTob:=Tob.Create('zz', nil, -1) ;
SCTob.LoadDetailDb('ECRITURE', '', '', Q, FALSE) ;
Ferme(Q) ;
TheData:=SCTob ;
AGLLanceFiche('CP', 'SUIVICLIENT', '', '', '') ;
SCTob.Free ;
TheData:=nil ;
end ;

{ TOF_SUIVICLIENT }

procedure TOF_SUIVICLIENT.CompteOnExit(Sender: TObject);
begin
     if (THEdit(Sender)=gen1) or (THEdit(Sender)=gen2) then
          DoCompteOnExit(THEdit(Sender), gen1, gen2);
     if (THEdit(Sender)=aux1) or (THEdit(Sender)=aux2) then
          DoCompteOnExit(THEdit(Sender), aux1, aux2);
end;

procedure TOF_SUIVICLIENT.DateOnExit(Sender: TObject);
begin
     DoDateOnExit(THEdit(Sender), Ech1, Ech2, DateD, DateF);
end;

procedure TOF_SUIVICLIENT.ExoOnChange(Sender: TObject);
begin
  DoExoToDateOnChange(Exo, Ech1, Ech2);
  DateD := StrToDate(Ech1.Text);
  DateF := StrToDate(Ech2.Text);
end;

procedure TOF_SUIVICLIENT.JalOnExit(Sender: TObject);
begin
     DoJalOnExit(THEdit(Sender), Jal1, Jal2);
end;

procedure TOF_SUIVICLIENT.SensOnChange(Sender: TObject);
begin
  if sens.Value = 'D' then
     xxWhere.Text := 'E_DEBIT<>0'
  else if sens.Value = 'C' then
     xxWhere.Text := 'E_CREDIT<>0'
  else
     xxWhere.Text := '';
end;

procedure TOF_SUIVICLIENT.TidcOnClick(Sender: TObject) ;
begin
if Tidc.Checked then
  begin
  gen1.DataType:='TZGTIDTIC' ;    gen2.DataType:='TZGTIDTIC' ;
  gen1.Text:='' ;                 gen2.Text:='' ;
  aux1.Enabled:=FALSE ;           aux2.Enabled:=FALSE ;
  aux1.Text:='' ;                 aux2.Text:='' ;
  end else
  begin
  gen1.DataType:='TZGCOLLECTIF' ; gen2.DataType:='TZGCOLLECTIF' ;
  gen1.Text:='' ;                 gen2.Text:='' ;
  aux1.Enabled:=TRUE ;            aux2.Enabled:=TRUE ;
  end ;
end ;

procedure TOF_SUIVICLIENT.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

procedure TOF_SUIVICLIENT.OnArgument(Arguments: string);
var Agrandir: TToolbarButton97 ;
begin
inherited ;
if TheData<>nil then
  begin
  TFStat(Ecran).CritereVisible:=FALSE ;
  TFStat(Ecran).ModeAlimentation:=maTOB ;
  TFStat(Ecran).LaTOB:=TOB(TheData) ;
  Agrandir:=TToolbarButton97(GetControl('BAGRANDIR')) ;
  if Agrandir<>nil then Agrandir.Visible:=FALSE ;
  end else
  begin
  with TForm(Ecran) do begin
    gen1:=THEdit(GetControl('E_GENERAL'));
    gen2:=THEdit(GetControl('E_GENERAL_'));
    aux1:=THEdit(GetControl('E_AUXILIAIRE'));
    aux2:=THEdit(GetControl('E_AUXILIAIRE_'));
    Jal1:=THEdit(GetControl('E_JOURNAL'));
    Jal2:=THEdit(GetControl('E_JOURNAL_'));
    Paie:=THValComboBox(GetControl('E_MODEPAIE')) ;
    Exo:=THValComboBox(GetControl('E_EXERCICE')) ;
    Etab:=THValComboBox(GetControl('E_ETABLISSEMENT')) ;
    Devise:=THValComboBox(GetControl('E_DEVISE')) ;
    Sens:=THValComboBox(GetControl('SENS')) ;
    Ech1:=THEdit(GetControl('E_DATECOMPTABLE'));  // 13724 E_DATEECHEANCE'));
    Ech2:=THEdit(GetControl('E_DATECOMPTABLE_')); // 13724 E_DATEECHEANCE_'));
    xxwhere:=THEdit(GetControl('XX_WHERE'));
    Tidc:=TCheckBox(GetControl('CBTICTID')) ;
    Rib:=TCheckBox(GetControl('CBRIB')) ;
    PzLibre:=TTabSheet(GetControl('PZLIBRE')) ;
    pages := TPageControl(GetControl('Pages'));
    TViewer:=TTobViewer(GetControl('TV')) ;
    InitTous;
  end;
  LibellesTableLibre(PzLibre, 'TT_TABLE', 'T_TABLE', 'T') ;
  end ;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    Aux1.OnElipsisClick:=AuxiElipsisClick;
    Aux2.OnElipsisClick:=AuxiElipsisClick;
  end;
end;

procedure TOF_SUIVICLIENT.OnUpdate ;
begin
inherited ;
//if TheData<>nil then Exit ;
end ;

procedure TOF_SUIVICLIENT.OnNew ;
begin
inherited ;
{ DblClick }
if TViewer=nil then Exit ;
ShowMessage('Row=' + IntToStr( TViewer.CurrentRow ) + ' Col=' +IntToStr( TViewer.CurrentCol ));
end ;

procedure TOF_SUIVICLIENT.OnLoad ;
var where: string ;
begin
if TheData<>nil then Exit ;
//if TheData<>nil then begin TFStat(Ecran).AutoSearch:=asMouette ; Exit ; end ;
inherited ;
if xxwhere<>nil then
  begin
  where:='E_QUALIFPIECE="N" AND E_ECRANOUVEAU<>"OAN"' ;
  if VH^.ExoV8.Code='' then where:=where+' AND E_ECRANOUVEAU<>"H"' ;
  if (Rib<>nil) and (Rib.State=cbChecked)   then where:=where+' AND E_RIB<>""' ;
  if (Rib<>nil) and (Rib.State=cbUnchecked) then where:=where+' AND E_RIB=""' ;
  SetControlText('XX_WHERE', where) ;
  end ;
end ;

procedure TOF_SUIVICLIENT.InitTous;
begin
  if gen1 <> nil then gen1.OnExit := CompteOnExit;
  if gen2 <> nil then gen2.OnExit := CompteOnExit;
  if aux1 <> nil then aux1.OnExit := CompteOnExit;
  if aux2 <> nil then aux2.OnExit := CompteOnExit;
  if Jal1 <> nil then Jal1.OnExit := JalOnExit;
  if Jal2 <> nil then Jal2.OnExit := JalOnExit;
  if Ech1 <> nil then Ech1.OnExit := DateOnExit ;
  if Ech2 <> nil then Ech2.OnExit := DateOnExit ;
  if Exo  <> nil then begin
     Exo.OnChange:=ExoOnChange;
     Exo.Value:=VH^.Entree.Code;
  end;
  if Tidc<>nil then Tidc.OnClick:=TidcOnClick ;
  if Paie <> nil then
     Paie.ItemIndex := 0;
  if (Etab<>nil) and (Etab.ItemIndex=-1) then Etab.ItemIndex := 0;
  if Devise <> nil then
     Devise.ItemIndex := 0;
  if sens <> nil then begin
     sens.OnChange := SensOnChange;
     sens.value := 'M';
  end;
//  Pages.Pages[2].TabVisible := false;
  Pages.ActivePage:=Pages.Pages[0];
  xxWhere.text := '';

end;

initialization
TheData:=nil ;
RegisterClasses([TOF_SUIVICLIENT]) ;


end.
