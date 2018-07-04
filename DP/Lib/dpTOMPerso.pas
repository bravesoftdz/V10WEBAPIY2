unit dpTOMPerso;
// TOM de la table DPPERSO

interface
uses StdCtrls, Controls, Classes, db, forms, sysutils, ComCtrls,
  HCtrls, HEnt1, HMsgBox, UTOM, UTOB;
type
  TOM_DPPERSO = class(TOM)
    procedure OnArgument(stArgument: string); override;
    procedure OnLoadRecord; override;

  private
    bCharge  :boolean; // pour éviter les OnChange pendant les onload

    procedure ChangeAdresse;
    procedure DPP_VOIENO_OnExit(Sender: TObject);
    procedure DPP_VOIETYPE_OnExit(Sender: TObject);
    procedure DPP_VOIENOCOMPL_OnExit(Sender: TObject);
    procedure DPP_VOIENOM_OnExit(Sender: TObject);
  end;


procedure DPP_ChangeAdresse (adr_:tob) ; //LM20070213

 ///////////// IMPLEMENTATION ////////////
implementation

uses dpOutils, DpJurOutils;

procedure TOM_DPPERSO.OnArgument(stArgument: string);
begin
  inherited;

  // 20/06/01
  SetControlVisible('BINSERT', False);

  // $$$ JP 28/09/05 - gestion saisie adresse
  THEdit (GetControl ('DPP_VOIENO')).OnExit := DPP_VOIENO_OnExit;
  THEdit (GetControl ('DPP_VOIENOM')).OnExit := DPP_VOIENOM_OnExit;
  THValComboBox (GetControl('DPP_VOIETYPE')).OnExit := DPP_VOIETYPE_OnExit;
  THValComboBox (GetControl('DPP_VOIENOCOMPL')).OnExit := DPP_VOIENOCOMPL_OnExit;
end;

procedure TOM_DPPERSO.OnLoadRecord;
var TOBAnnu: TOB;
  visib: Boolean;
begin
  bCharge := true;
  TOBAnnu := TOB.Create('ANNUAIRE', nil, -1);
  try
     TOBAnnu.SelectDB('"'+GetField('DPP_GUIDPER')+'"', nil, TRUE);

     // valeur de ANN_SITUFAM basé sur tablette JUSITUFAM :
     // si CELibataires ou INConnu (non listé dans la tablette),
     // on ne saisit pas le régime matrimonial DPP_REGIMEMATRI (tablette DPREGIMMATRI)
     // (CA/CR/CU ... Communauté réduite aux acquets, ...)
     // $$$ JP 29/09/05 - condition supplémentaire pour ne pas voir le régimé matri, ni onglet conjoint: si personne morale
     visib := (TOBAnnu.GetValue('ANN_SITUFAM') <> 'CEL') and (TOBAnnu.GetValue('ANN_SITUFAM') <> 'INC') and (TOBAnnu.GetString ('ANN_PPPM') <> 'PM');
     SetControlVisible('DPP_REGIMEMATRI', visib);
     SetControlVisible('TDPP_REGIMEMATRI', visib);
     TTabSheet(GetControl('TSCJ')).TabVisible := visib; // onglet conjoint

     // MD 1209/06 - Modif libellé pour CJ
     if TOBAnnu.GetString ('ANN_PPPM')='PP' then
          SetControlCaption('DPP_EXPLOITDOM', 'Gestion d''une adresse personnelle')
     else
          SetControlCaption('DPP_EXPLOITDOM', 'Gestion d''une adresse fiscale');

     // à la 1ère création, reprend les coord. de la fiche annuaire
     if (DS <> nil) and (DS.State = dsInsert) then
     begin
          SetField('DPP_RESIDENCE', TOBAnnu.GetValue('ANN_ALRESID')); // 30c <= 17c
          SetField('DPP_BATIMENT', TOBAnnu.GetValue('ANN_ALBAT')); // 5c  <=  4c
          SetField('DPP_ESCALIER', TOBAnnu.GetValue('ANN_ALESC')); // 5c  <=  4c
          SetField('DPP_ETAGE', TOBAnnu.GetValue('ANN_ALETA')); // 5c  <=  4c
          SetField('DPP_APPARTEMENT', TOBAnnu.GetValue('ANN_ALNOAPP')); // 5c <= 4c
          SetField('DPP_ADRESSE1', TOBAnnu.GetValue('ANN_ALRUE1')); // 35c <= 35c
          SetField('DPP_ADRESSE2', TOBAnnu.GetValue('ANN_ALRUE2')); // idem
          SetField('DPP_ADRESSE3', TOBAnnu.GetValue('ANN_ALRUE3')); // idem
          SetField('DPP_CODEPOSTAL', TOBAnnu.GetValue('ANN_ALCP')); // 9c <= 9c
          SetField('DPP_VILLE', TOBAnnu.GetValue('ANN_ALVILLE')); // 35c <= 35c
          SetField('DPP_PAYS', TOBAnnu.GetValue('ANN_PAYS')); // 3c <= 3c

          //CM 15/06/2005 : ajout des nouveaux champs
          SetField('DPP_VOIENO', TOBAnnu.GetValue('ANN_VOIENO'));
          SetField('DPP_VOIENOCOMPL', TOBAnnu.GetValue('ANN_VOIENOCOMPL'));
          SetField('DPP_VOIETYPE', TOBAnnu.GetValue('ANN_VOIETYPE'));
          SetField('DPP_VOIENOM', TOBAnnu.GetValue('ANN_VOIENOM'));
     end;
  finally
         FreeAndNil (TOBAnnu);
  end;

  TOBAnnu := TOB.Create('ANNUBIS', nil, -1);
  try
     TOBAnnu.SelectDB('"'+GetField('DPP_GUIDPER')+'"', nil, TRUE);

     // onglet fiscal
     visib := (TOBAnnu.GetValue('ANB_GESTIONFISC') = 'X');
     TTabSheet(GetControl('TSFISC')).TabVisible := visib;

     // onglet social
     visib := (TOBAnnu.GetValue('ANB_COTISETNS') = 'X');
     TTabSheet(GetControl('TSSOC')).TabVisible := visib;

     // récup de valeurs dans la table ANNUBIS, car ce ne sont pas des zones affichées dans la fiche FICHPERSONNE
     if GetField('DPP_GESTIONFISC') <> TOBAnnu.GetValue('ANB_GESTIONFISC') then
     begin
          ModeEdition(DS);
          SetField('DPP_GESTIONFISC', TOBAnnu.GetValue('ANB_GESTIONFISC'));
     end;
     if GetField('DPP_GESTIONPATRIM') <> TOBAnnu.GetValue('ANB_GESTIONPATRIM') then
     begin
          ModeEdition(DS);
          SetField('DPP_GESTIONPATRIM', TOBAnnu.GetValue('ANB_GESTIONPATRIM'));
     end;
     if GetField('DPP_COTISETNS') <> TOBAnnu.GetValue('ANB_COTISETNS') then
     begin
          ModeEdition(DS);
          SetField('DPP_COTISETNS', TOBAnnu.GetValue('ANB_COTISETNS'));
     end;
  finally
         FreeAndNil (TOBAnnu);
  end;

  // $$$ JP 17/10/2003: ne pas autoriser saisie adresse pro si pas cochée
  //LM20070411 SetControlVisible('GPBOX1', GetField('DPP_EXPLOITDOM') = 'X');

  //if TOBAnnu <> nil then TOBAnnu.Free;
  bCharge := false;
end;

procedure DPP_ChangeAdresse (adr_:tob) ; //LM20070213
var st:string;
begin
   if adr_.g('DPP_VOIENO') <> '0' then     st := adr_.g('DPP_VOIENO');
   if adr_.g('DPP_VOIENOCOMPL') <> '' then st := st + ' ' + RechDom('YYVOIENOCOMPL', adr_.g('DPP_VOIENOCOMPL'), TRUE);
   if adr_.g('DPP_VOIETYPE') <> '' then    st := st + ' ' + RechDom('YYVOIETYPE', adr_.g('DPP_VOIETYPE'), TRUE);
   if adr_.g('DPP_VOIENOM') <> '' then     st := st + ' ' + adr_.g('DPP_VOIENOM');
   st := Trim (st);
   adr_.p('DPP_ADRESSE1', st);
end ;

procedure TOM_DPPERSO.ChangeAdresse;
var adresse  :string;
    adr_:tob ;
begin
     if bCharge then exit;
     adr_:=tob.create('DPPERSO', nil, -1 );
     adr_.GetEcran(Ecran);
     DPP_changeAdresse(adr_) ;
     if adr_.g('DPP_ADRESSE1')<> GetField ('DPP_ADRESSE1') then
      SetField ('DPP_ADRESSE1', adr_.g('DPP_ADRESSE1'));
     adr_.free ;
     (*
     adresse := '';
     if GetControlText('DPP_VOIENO') <> '0' then
        adresse := GetControlText('DPP_VOIENO');
     if GetControlText('DPP_VOIENOCOMPL') <> '' then
        adresse := adresse + ' ' + RechDom('YYVOIENOCOMPL', GetControlText('DPP_VOIENOCOMPL'), TRUE);
     if GetControlText('DPP_VOIETYPE') <> '' then
        adresse := adresse + ' ' + RechDom('YYVOIETYPE', GetControlText('DPP_VOIETYPE'), TRUE);
     if GetControlText('DPP_VOIENOM') <> '' then
        adresse := adresse + ' ' + GetControlText('DPP_VOIENOM');
     adresse := Trim (adresse);
     if GetField ('DPP_ADRESSE1') <> adresse then
        SetField ('DPP_ADRESSE1', adresse);
     *)
end;

procedure TOM_DPPERSO.DPP_VOIENO_OnExit(Sender: TObject);
begin
  ChangeAdresse;
end;

procedure TOM_DPPERSO.DPP_VOIETYPE_OnExit(Sender: TObject);
begin
  ChangeAdresse;
end;

procedure TOM_DPPERSO.DPP_VOIENOCOMPL_OnExit(Sender: TObject);
begin
  ChangeAdresse;
end;

procedure TOM_DPPERSO.DPP_VOIENOM_OnExit(Sender: TObject);
begin
  ChangeAdresse;
end;


initialization
  registerclasses([TOM_DPPERSO]);

end.

