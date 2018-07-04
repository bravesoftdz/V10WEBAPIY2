{***********UNITE*************************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 24/01/2007
Modifié le ... :   /  /
Description .. : TOM Générale de Paie pour journal événement FQ 12694
Mots clefs ... : REMUNERATION,COTISATION,PROFIL,VARIABLE,CUMUL
*****************************************************************}
{
PT1   : 27/02/2007  GGS V800  Rajout test que nom du controle commence par préfixe
PT2   : 02/03/2007  GGS V800  Rajout Eléments nationaux dossier
PT3   : 05/11/2007  FC  V_80  FQ 14908 Rajout Elément dynamique, Table dynamique, Niveau préconisé, Cumul personnalisé
PT4   : 03/01/2008  FC  V_810 FQ 14762 Rajout fiche salarié
PT5   : 17/01/2008  GGU V_810 Correction d'un plantage (erreur de conversion) dans certains cas (par exemple pour les contrôles non visibles.
PT6   : 24/01/2008  FC  V_81  FQ 14149 Temporairement on ne gère pas la fiche salarié car bug agl sur fonction IsFieldModified FQ AGL 14796
}
unit PAIETOM;
interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
{$IFDEF EAGLCLIENT}
  UtileAGL, eFiche, MaineAgl,
{$ELSE}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, Fiche, DBCtrls, FE_Main, EdtREtat,
{$ENDIF$}
  HCtrls, HEnt1, HMsgBox, UTOM, UTOB, HTB97,
  P5Def,
  EntPaie,
  uTOMComm,wcommuns;


type
  PGTOM = class(TOM)
  protected
    EnDupl: String;
    function Isdifferent(creamod, par, par2, par3:String;LaTrace:TStringList;LaFiche:THForm; LeStatut: TDataSetState = DsBrowse):boolean;
  end;

implementation

function PGTOM.IsDifferent(creamod,Par,Par2,PAR3:String;LaTrace:TStringList;LaFiche:THForm; LeStatut: TDataSetState = DsBrowse):boolean;
var
    Edititre: boolean;
    Titre,Titrelabel,Ch10,CH11,CH20,CH21,LaTablette: String;
    Lecontrol: TComponent;
    i: Integer;
    Lelabel: THLabel;
    NomControl,NomLibelle,TypeEven,LibEven: string;
begin;
  TypeEven := '003';
  if par <> '' then
  begin
    Titre := par;
    NomControl := par2;
    Nomlibelle := par3;
    if par = 'REMUNERATION' then LibEven := '080';
    if par = 'COTISATION' then LibEven := '081';
    if par = 'PROFILPAIE' then LibEven := '082';
    if par = 'VARIABLEPAIE' then LibEven := '083';
    if par = 'CUMULPAIE' then LibEven := '084';
    if par = 'ELTNATIONAUX' then LibEven := '085';
    if par = 'PGEXCEPTIONS' then LibEven := '086';
    if par = 'MOTIFABSENCE' then LibEven := '087';
    if par = 'VENTICOTPAIE' then LibEven := '088';
    if par = 'VENTIREMPAIE' then LibEven := '089';
    if par = 'VENTIORGPAIE' then LibEven := '090';
    if par = 'DUCSAFFECT' then LibEven := '091';
    //DEB PT3
    if par = 'ELTNATIONDOS' then LibEven := '092';
    if par = 'PARAMSALARIE' then LibEven := '093';
    if par = 'TABLEDIMENT' then LibEven := '094';
    if par = 'ELTNIVEAUREQUIS' then LibEven := '095';
    if par = 'PGCUMEXCEPTIONS' then
    begin
      LibEven := '096';
      par := 'PGEXCEPTIONS';
    end;
    if par = 'SALARIES' then LibEven := '097';
    //FIN PT3
  end;
  if EnDupl='OUI' then
  begin
    CreeJnalEvt(TypeEven,LibEven,'OK',nil,nil,LaTrace);
    exit;
  end;
  if creamod = GetField(NomControl) then
  begin
    if EnDupl = 'NON' then titre := 'Modification suite duplication '+ Par+' '
    else
    begin
     if LeStatut = DsBrowse then titre := 'Modification '+ Par+' '
     else titre := 'Création '+ Par+' ';
    end;
    titre := titre  + GetField(NomControl)+' '+ GetField(NomLibelle);
    LaTrace.Add(Titre);
    CreeJnalEvt(TypeEven,LibEven,'OK',nil,nil,LaTrace);
  end
  else
  begin
    edititre := True;
    titre := 'Modification '+par+' '+GetField(NomControl);
    for i := 0 to Lafiche.ComponentCount - 1 do
    begin
      TitreLabel := '';
      LeControl := Lafiche.Components[i];
      if pos('_',LeControl.Name) > 0 then         //PT1
      begin
        if (LeControl is TCustomEdit) or
           (LeControl is THValComboBox) or (LeControl is TCustomCheckBox) or
           ( LeControl is THEdit) then
        begin
          //PT3 Tester l'existence du champ avant de tester qu'il a été modifié car certains controles ont un préfixe
          //mais ne sont pas des champs (cf TABLEDIMENT)
          //if IsFieldModified(LeControl.Name) then
          if (par <> 'SALARIES') and wExistFieldInDechamps(LeControl.Name) and IsFieldModified(LeControl.Name) then   //PT3   //PT6
          begin
            LeLabel := THLabel(GetControl ('T' +  LeControl.name));
            if (LeControl is THValComboBox) then
            begin
              LaTablette := THValComboBox(LeControl).DataType;
              Ch10 := GetFieldAvantModif(LeControl.Name);
              ch11 := RechDom (LaTablette,GetFieldAvantModif(LeControl.Name),FALSE);
              try //PT5
                Ch20 := GetField(LeControl.Name);
              except
                Ch20 := GetControlText(LeControl.Name);
              end;
              ch21 := RechDom (LaTablette,Ch20,FALSE);
            end
            else
            begin
              if (LeControl is TCustomEdit) or (LeControl is THEdit) then
              begin
                Ch10 := '';
                Ch11 := GetFieldAvantModif(LeControl.Name);
                Ch20 := '';
                try //PT5
                  Ch21 := GetField(LeControl.Name);
                except
                  Ch21 := GetControlText(LeControl.Name);
                end;
              end
              else
              begin
                if (LeControl is TCheckBox) or (LeControl is TCustomCheckBox) then
                begin
                  try //PT5
                    Ch21 := GetField(LeControl.Name);
                  except
                    Ch21 := GetControlText(LeControl.Name);
                  end;
                  if Ch21 = 'X' then
                  begin
                    Ch10:= '';
                    Ch11:='';
                    Ch20 := '';
                    Ch21 := 'Coché';
                  end
                  else
                  begin
                    Ch10:= '';
                    Ch11:='';
                    Ch20 := '';
                    Ch21 := 'Décoché';
                  end;
                end
                else
                begin
                  Ch10 := '';
                  Ch11 := '';
                  Ch20 := '';
                  ch21 := '';
                end;
              end;
            end;
            if edititre then
            begin
              LaTrace.Add(Titre);
              edititre := False;
            end;
            if LeLabel <> NIL then TitreLabel := THLabel(LeLabel).caption
            else
            begin
            if (LeControl is TCustomCheckBox) then TitreLabel := THLabel(LeControl).caption;
            end;
            if Titrelabel = '' then Titrelabel := '???';
            LaTrace.Add(TitreLabel+' '+Ch10+' '+Ch11+' devient '+Ch20+' '+ch21);
            titre := '';
          end;
        end;
      end;
    end;
  end;
  if titre = '' then CreeJnalEvt(TypeEven,LibEven,'OK',nil,nil,LaTrace);
end;
initialization
  registerclasses([PGTOM]);
end.

