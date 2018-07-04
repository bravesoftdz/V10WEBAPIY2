unit UtofProspectConf;

interface
uses HTB97,HEnt1, hmsgbox, ExtCtrls,
{$IFDEF EAGLCLIENT}
     MaineAGL,eMul,
{$ELSE}
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}FE_Main,Mul,
{$ENDIF}
     Hctrls,
     StdCtrls, Forms,Controls,
     UTOF, Classes,Graphics,SysUtils,comctrls,Vierge,M3FP,
     UTOB,HQry,EntRT;

type
     TOF_RTMULConfident = Class(TOF)
       procedure OnArgument (S : String ) ; override ;
     private
       procedure ChargeMulIntervenant(Intervenant:string);
     end;

     TOF_RTConfident = Class (TOF)
     private
         Intervenant,DroitProfil,DroitIntervenant,stProduitpgi :string;
         bTriLibelle : boolean;
         TobAcces,Tobconsult,TobModif,Tobconsult2,TobModif2,TobAction : TOB;
         TobCreatPro,TobCreatCon,TobCreatInfos,TobCreatAction : TOB;
         PCons : TPanel;
         CBIntervenant: THValcomboBox;
         CBProfil: THValcomboBox;
         CBCHAMP1 : THValcomboBox;
         CBCHAMP2 : THValcomboBox;
         CBCHAMP3 : THValcomboBox;
         CBCHAMP1_ : THValcomboBox;
         CBCHAMP2_ : THValcomboBox;
         CBCHAMP3_ : THValcomboBox;
         CBCHAMP1__ : THValcomboBox;
         CBCHAMP2__ : THValcomboBox;
         CBCHAMP3__ : THValcomboBox;
         CBCHAMP1___ : THValcomboBox;
         CBCHAMP2___ : THValcomboBox;
         CBCHAMP3___ : THValcomboBox;
         LBACT,LBACTSEL,LBACTCODESEL: TListBox;
         LACTCODE:TStrings;
         ModifLot,ResponsableAct,AutoriseAct :Boolean;
         MemoCons,MemoMod:Tmemo;
         procedure BINTERVENANTClick (Sender: TObject);
         procedure BPROFILClick (Sender: TObject);
         procedure BSQLClick (Sender: TObject);
         procedure BTOUTACCESClick (Sender: TObject);
         procedure BCOPIEACCESClick (Sender: TObject);
         procedure BCHAMP1Click (Sender: TObject);
         procedure BCHAMP2Click (Sender: TObject);
         procedure BCHAMP3Click (Sender: TObject);
         procedure BCHAMP1_Click (Sender: TObject);
         procedure BCHAMP2_Click (Sender: TObject);
         procedure BCHAMP3_Click (Sender: TObject);
         procedure BCHAMP1__Click (Sender: TObject);
         procedure BCHAMP2__Click (Sender: TObject);
         procedure BCHAMP3__Click (Sender: TObject);
         procedure BCHAMP1___Click (Sender: TObject);
         procedure BCHAMP2___Click (Sender: TObject);
         procedure BCHAMP3___Click (Sender: TObject);
         function  ControleCritere (TypeConf: String;TobConf:TOB):string;
         procedure ControleTousCritere;
         function  AffecteDatatype(CHAMPTABLE,CHAMPCONF: string):string;
         function  ChargeCritere (TypeConf: String):string;
         procedure AfficheCritere (TypeConf: String);
         procedure ChargeIntervenant;
         function  GetValue (Champtable: String):string;
         procedure PutValue(Champtable:string; Valeur: Variant);
         procedure BAJOUTERClick (Sender: TObject);
         procedure BENLEVERClick (Sender: TObject);
         procedure EnregistreIntervenant(count : integer;Maj:boolean);
         procedure SignaleTTDroitModif(profil : boolean);
         function  SupprimeToutacces:boolean;
         function  DonneToutacces :boolean;
         procedure ApercuRequete;
         procedure CopieAccesCons;
         procedure chargeComboChamps( Prefixe,ListeDhControle : string;ListeChamp,ListeLibelle: TStrings);
         procedure deleteIntervenant;
         function ModifieLesAcces:Boolean;
         procedure OnCreatActionClick (Sender: TObject);
         procedure GriseModAction (coche : boolean);
      public
         procedure OnArgument(Arguments : string) ; override ;
         procedure OnLoad ; override ;
         procedure OnUpdate ; override ;
         procedure OnClose ; override ;
         procedure OnDelete ; override ;
     Private
     END;

var stWhere,StWhereCons,StWhereMod,StWhereCons2,StWhereMod2,stWhereAct,stWhereConsAvant,stWhereCons2Avant,stWhereCon,stWherePro,stWhereCic,stWhereAction: string ;
    ModifLot : Boolean;
    ListeIntervenant : TStrings;

const
 FMess: Array[0..3] of string =
  {0}('0;?caption?;Confirmez-vous la suppression des droits d''accès?;Q;YN;Y;N;'
  {1},'0;?caption?;Confirmez-vous de donner tous les droits d''accès?;Q;YN;Y;N;'
  {2},'0;?caption?; Attention : vous avez modifié les droits d''accès en consultation :#13#10 êtes-vous sûrs de la cohérence avec les droits d''accès en création/modification ?;Q;YN;Y;N;'
  {3},'0;?caption?;Confirmez-vous la modification des droits d''accès?;Q;YN;Y;N;'
  ) ;
  
implementation
uses 
   CbpMCD
   ,CbpEnumerator
;

{ TOF_RTConfident }
function TOF_RTConfident.AffecteDatatype(CHAMPTABLE,CHAMPCONF: string):string;
var datatype :string;
begin
   result := '';
   if  (CHAMPTABLE = '') then datatype := ''
   else if copy(CHAMPTABLE,1,19)='YTC_TABLELIBRETIERS' then    //YTC_TABLELIBRETIERS0 -> GCLIBRETIERS0
      datatype := FindEtReplace(CHAMPTABLE,'YTC_TABLE','GC',false)
   else if copy(CHAMPTABLE,1,15)='RPR_RPRLIBTABLE' then   //RPR_RPRLIBTABLE0 -> RT RPRLIBTABLE0
      datatype := FindEtReplace(CHAMPTABLE,'RPR_','RT',false)
   else if (CHAMPTABLE='T_REPRESENTANT') or (CHAMPTABLE='YTC_REPRESENTANT2') or (CHAMPTABLE='YTC_REPRESENTANT3') then   
      datatype := 'GCCOMMERCIAL'
   else
      datatype := Get_Join(CHAMPTABLE );
   if (CHAMPCONF <> '') then
   begin
     if (THValcomboBox(GetControl(CHAMPCONF)).datatype <> datatype) then
     begin
       THValcomboBox(GetControl(CHAMPCONF)).values.clear;
       THValcomboBox(GetControl(CHAMPCONF)).items.clear;
       THValcomboBox(GetControl(CHAMPCONF)).text := '';
       THValcomboBox(GetControl(CHAMPCONF)).datatype := datatype;
     end;
     if  (CHAMPTABLE = '') then
     begin
       THValcomboBox(GetControl(FindEtReplace(CHAMPCONF,'RTC_VAL','RTC_OPER',false))).itemindex := -1;
       THValcomboBox(GetControl(CHAMPCONF)).text := '';
     end;
   end else result := datatype ;
end;

procedure TOF_RTConfident.BCHAMP1Click(Sender: TObject);
var ChampLibre :  string;
begin
   ChampLibre := THValcomboBox(GetControl('RTC_CHAMP1')).value;
   AffecteDatatype(ChampLibre,'RTC_VAL1');
end;

procedure TOF_RTConfident.BCHAMP2Click(Sender: TObject);
var ChampLibre : string;
begin
   ChampLibre := THValcomboBox(GetControl('RTC_CHAMP2')).value;
   AffecteDatatype(ChampLibre,'RTC_VAL2');
end;

procedure TOF_RTConfident.BCHAMP3Click(Sender: TObject);
var ChampLibre : string;
begin
   ChampLibre := THValcomboBox(GetControl('RTC_CHAMP3')).value;
   AffecteDatatype(ChampLibre,'RTC_VAL3');
end;

procedure TOF_RTConfident.BCHAMP1_Click(Sender: TObject);
var ChampLibre : string;
begin
   ChampLibre := THValcomboBox(GetControl('RTC_CHAMP1_')).value;
   AffecteDatatype(ChampLibre,'RTC_VAL1_');
end;

procedure TOF_RTConfident.BCHAMP2_Click(Sender: TObject);
var ChampLibre : string;
begin
   ChampLibre := THValcomboBox(GetControl('RTC_CHAMP2_')).value;
   AffecteDatatype(ChampLibre,'RTC_VAL2_');
end;

procedure TOF_RTConfident.BCHAMP3_Click(Sender: TObject);
var ChampLibre : string;
begin
   ChampLibre := THValcomboBox(GetControl('RTC_CHAMP3_')).value;
   AffecteDatatype(ChampLibre,'RTC_VAL3_');
end;

procedure TOF_RTConfident.BCHAMP1__Click(Sender: TObject);
var ChampLibre : string;
begin
   ChampLibre := THValcomboBox(GetControl('RTC_CHAMP1__')).value;
   AffecteDatatype(ChampLibre,'RTC_VAL1__');
end;

procedure TOF_RTConfident.BCHAMP2__Click(Sender: TObject);
var ChampLibre : string;
begin
   ChampLibre := THValcomboBox(GetControl('RTC_CHAMP2__')).value;
   AffecteDatatype(ChampLibre,'RTC_VAL2__');
end;

procedure TOF_RTConfident.BCHAMP3__Click(Sender: TObject);
var ChampLibre : string;
begin
   ChampLibre := THValcomboBox(GetControl('RTC_CHAMP3__')).value;
   AffecteDatatype(ChampLibre,'RTC_VAL3__');
end;

procedure TOF_RTConfident.BCHAMP1___Click(Sender: TObject);
var ChampLibre : string;
begin
   ChampLibre := THValcomboBox(GetControl('RTC_CHAMP1___')).value;
   AffecteDatatype(ChampLibre,'RTC_VAL1___');
end;

procedure TOF_RTConfident.BCHAMP2___Click(Sender: TObject);
var ChampLibre : string;
begin
   ChampLibre := THValcomboBox(GetControl('RTC_CHAMP2___')).value;
   AffecteDatatype(ChampLibre,'RTC_VAL2___');
end;

procedure TOF_RTConfident.BCHAMP3___Click(Sender: TObject);
var ChampLibre : string;
begin
   ChampLibre := THValcomboBox(GetControl('RTC_CHAMP3___')).value;
   AffecteDatatype(ChampLibre,'RTC_VAL3___');
end;

procedure TOF_RTConfident.BAJOUTERClick(Sender: TObject);
var i_ind1 : integer;
    st_trav1 : string;
begin
  for i_ind1 := 0 to LBACT.Items.Count - 1 do
  if LBACT.Selected[i_ind1] then
  begin
    st_trav1 := LBACT.Items.Strings[i_ind1];
    LBACTSEL.Items.Append(st_trav1);
    LBACTCODESEL.Items.Append(LACTCODE[i_ind1]);
  end;
end;

procedure TOF_RTConfident.BENLEVERClick(Sender: TObject);
var Index : integer;
begin
  for Index := LBACTSEL.items.count - 1 downto 0 do
  if LBACTSEL.Selected[Index] then
  begin
    LBACTSEL.items.Delete (Index);
    LBACTCODESEL.Items.Delete (Index);
  end;
end;

function VireLeWhere(stWhere: string) : string;
var i : integer;
begin
  Result := trim (stWhere);
  if stWhere<>'' then
  begin
    //i:=pos('WHERE (AND',stWhere);
    //if i>0 then Result:=' AND '+copy(stwhere,i+10,length(stwhere)-12);
    i:=pos('WHERE ((',stWhere);
    if i>0 then Result:=' AND ('+copy(stwhere,i+8,length(stwhere)-11)+')';
  end;
end;

procedure TOF_RTConfident.BSQLClick(Sender: TObject);
begin
  ControleTousCritere;
  ApercuRequete;
end;

procedure TOF_RTConfident.ApercuRequete;
var stTab :string;
begin
  MemoCons.Clear; MemoMod.Clear;
  stTab := '                                        ';
  if (length(stWhereCons)<=1) and (length(stWhereCons2)<=1) then
    MemoCons.Lines.Text := stTab+TraduireMemoire('Accès en consultation à toutes les fiches tiers')
  else
    MemoCons.Lines.Text := stTab+TraduireMemoire('Consultation Fiche Tiers Critère 1');

  MemoMod.Lines.Text  := stTab+TraduireMemoire('Création/Modification Fiche Tiers Critère 1');
  if length(stWhereCons)>1 then MemoCons.Lines.Append(stWhereCons);
  if length(stWhereCons2)>1 then
  begin
    MemoCons.Lines.Append(stTab+TraduireMemoire('Ou critère 2'));
    MemoCons.Lines.Append(stWhereCons2);
  end;

  if (stProduitpgi = 'GRC' ) then
    begin
    if ( TCheckBox(GetControl('CREATPROPO')).checked ) then
       MemoCons.Lines.Append(TraduireMemoire('Création de propositions étendue à tous les tiers'))
    else
       MemoCons.Lines.Append(TraduireMemoire('Création de propositions limitée aux tiers accessibles en modification'));
    if ( TCheckBox(GetControl('CREATACTION')).checked ) then
       MemoCons.Lines.Append(TraduireMemoire('Création d''actions étendue à tous les tiers'))
    else
       MemoCons.Lines.Append(TraduireMemoire('Création d''actions limitée aux tiers accessibles en modification'));

    end;

  if ( TCheckBox(GetControl('CREATCONTACT')).checked ) then
     MemoCons.Lines.Append(TraduireMemoire('Création de contacts étendue à tous les tiers'))
  else
     MemoCons.Lines.Append(TraduireMemoire('Création de contacts limitée aux tiers accessibles en modification'));

  if ( TCheckBox(GetControl('CREATINFOS')).checked ) then
     MemoCons.Lines.Append(TraduireMemoire('Création des infos complémentaires étendue à tous les tiers'))
  else
     MemoCons.Lines.Append(TraduireMemoire('Création des infos complémentaires limitée aux tiers accessibles en modification'));

  if length(stWhereMod)>1 then MemoMod.Lines.Append(stWhereMod);
  if length(stWhereMod2)>1 then
  begin
    MemoMod.Lines.Append(stTab+TraduireMemoire('Ou Critère 2'));
    MemoMod.Lines.Append(stWhereMod2);
  end;

  if (ResponsableAct) then MemoMod.Lines.Append(stTab+TraduireMemoire('Modification limitée à ses propres actions'));
  if length(stWhereAct)>1 then
  begin
    if AutoriseAct then MemoMod.Lines.Append(stTab+TraduireMemoire('Modification autorisée des types Actions'))
    else MemoMod.Lines.Append(stTab+TraduireMemoire('Modification interdite des types Actions'));
    MemoMod.Lines.Append(stWhereAct);
  end;
end;

procedure TOF_RTConfident.BINTERVENANTClick(Sender: TObject);
begin
  Intervenant := CBIntervenant.value;
  ChargeIntervenant;
  SignaleTTDroitModif(False);
  CBProfil.ItemIndex := 0;DroitProfil:=''; 
  THLABEL(Getcontrol('TTDROITPROFIL')).caption := '';
end;

procedure TOF_RTConfident.BPROFILClick(Sender: TObject);
begin
  Intervenant := CBProfil.value;
  if (Intervenant <> '') then
  begin
    ChargeIntervenant;
    SignaleTTDroitModif(True)
  end else
  begin
    THLABEL(Getcontrol('TTDROITPROFIL')).caption := '';DroitProfil:='';
  end
end;

function TOF_RTConfident.ChargeCritere (TypeConf: String):string;
var QQ :TQUERY;
    C : TControl ;
    i : integer;
begin
  Result := '-';
  TobAcces.InitValeurs(false);
  QQ := OpenSQL('Select * from PROSPECTCONF Where RTC_INTERVENANT="' + Intervenant + '" and RTC_TYPECONF="'+TypeConf+'" and RTC_PRODUITPGI="'+stProduitpgi+'"',True) ;
  if Not QQ.EOF then
  begin
    TobAcces.SelectDB('',QQ);
    AfficheCritere(TypeConf);
    Result := TobAcces.getvalue('RTC_SQLCONF');
  end else
  begin
    if (TypeConf='ACT') then
    begin
      LBACTCODESEL.clear;LBACTSEL.clear;
      THRadioGroup(GetControl('RGOPERATE')).Value:='=';
      TCheckBox(GetControl('CBCREATEUR')).checked := False;
      TCheckBox(GetControl('CBTOUSTYPES')).checked := False;
      ResponsableAct := False;
    end else
    if (TypeConf='CPR') then
    begin
      TCheckBox(GetControl('CREATPROPO')).checked := False;
    end else
    if (TypeConf='CAC') then
    begin
      TCheckBox(GetControl('CREATACTION')).checked := False;
    end else
    if (TypeConf='CCO') then
    begin
      TCheckBox(GetControl('CREATCONTACT')).checked := False;
    end else
    if (TypeConf='CIC') then
    begin
      TCheckBox(GetControl('CREATINFOS')).checked := False;
    end else
    begin
      PCons := TPanel(getcontrol('PTIERS'+TypeConf));
      TobAcces.PutEcran (TFVierge(Ecran),PCons);
      for i:=1 to PCons.ControlCount-1 do
      begin
        C:=PCons.Controls[i] ;
        if (Copy(C.Name,1,7) = 'RTC_VAL') then THValcomboBox(C).datatype := '';
      end;
    end;
  end;
  Ferme (QQ);
end;

procedure TOF_RTConfident.ChargeIntervenant;
begin
  StWhereCons  := ChargeCritere ('CON');
  StWhereCons2 := ChargeCritere ('CO2');
  StWhereMod   := ChargeCritere ('MOD');
  StWhereMod2  := ChargeCritere ('MO2');
  StWhereAct   := ChargeCritere ('ACT');
  StWherePro   := ChargeCritere ('CPR');
  StWhereCon   := ChargeCritere ('CCO');
  StWhereCic   := ChargeCritere ('CIC');
  stWhereAction:= ChargeCritere ('CAC');

  stWhereConsAvant := stWhereCons ;
  stWhereCons2Avant := stWhereCons2 ;
  TobAcces.putvalue('RTC_INTERVENANT',Intervenant);
  TobAcces.putvalue('RTC_PRODUITPGI',stProduitpgi);
  ApercuRequete;
end;

procedure TOF_RTConfident.AfficheCritere(TypeConf: String);
var C : TControl ;
    CBValeur : THValComboBox ;
    ChampLibre,CValeur,COper : string;
    i ,index: integer;
begin
    if (TypeConf = 'CPR') then
      begin
      TCheckBox(GetControl('CREATPROPO')).checked := (TobAcces.getvalue('RTC_VAL1') = '=');
      end
    else
    if (TypeConf = 'CAC') then
      begin
      TCheckBox(GetControl('CREATACTION')).checked := (TobAcces.getvalue('RTC_VAL1') = '=');
      end
    else
    if (TypeConf = 'CCO') then
      begin
      TCheckBox(GetControl('CREATCONTACT')).checked := (TobAcces.getvalue('RTC_VAL1') = '=');
      end
    else
    if (TypeConf = 'CIC') then
      begin
      TCheckBox(GetControl('CREATINFOS')).checked := (TobAcces.getvalue('RTC_VAL1') = '=');
      end
    else
    if (TypeConf = 'ACT') then
    begin
      LBACTCODESEL.clear;LBACTSEL.clear;
      CValeur := TobAcces.getvalue('RTC_SQLCONF');
      while (CValeur <> '') do
      begin
        ChampLibre := ReadTokenSt (CValeur);
        LBACTCODESEL.Items.Append(ChampLibre);
        if (stProduitpgi = 'GRC' ) then
           LBACTSEL.Items.Append(RechDom('RTTYPEACTION',ChampLibre,FALSE))
        else
           LBACTSEL.Items.Append(RechDom('RTTYPEACTIONFOU',ChampLibre,FALSE));
      end;
      COper := TobAcces.getvalue('RTC_OPER1');
      if (COper = '') then COper := '=';
      THRadioGroup(GetControl('RGOPERATE')).Value:=COper;
      AutoriseAct := (COper = '=');
      COper := TobAcces.getvalue('RTC_OPER2');
      TCheckBox(GetControl('CBCREATEUR')).checked := (COper = '=');
      TCheckBox(GetControl('CBTOUSTYPES')).checked := (TobAcces.getvalue('RTC_VAL1') = '=');
      ResponsableAct := (COper = '=');
    end else
    begin
      PCons := TPanel(getcontrol('PTIERS'+TypeConf));
      for i:=1 to PCons.ControlCount-1 do
      begin
         C:=PCons.Controls[i] ;
         CValeur := FindEtReplace (C.Name,'CHAMP','VAL',false);
         CBValeur := THValcomboBox(GetControl(CValeur));
         if Copy(C.Name,5,2)='CH' then
         begin
            ChampLibre := GetValue(C.Name);
            THValcomboBox(GetControl(C.Name)).value := ChampLibre;
            CBValeur.datatype := '';  CBValeur.text := '';
            CBValeur.values.clear;    CBValeur.items.clear;
            if (pos ('LIBRETIERS',ChampLibre)>0) or (ChampLibre='T_REPRESENTANT')
            or (ChampLibre='YTC_REPRESENTANT2') or (ChampLibre='YTC_REPRESENTANT3')
            or (ChampToType( ChampLibre) = 'COMBO') then
            begin
              CBValeur.datatype := AffecteDatatype(ChampLibre, '');
              index := CBValeur.Values.IndexOf(GetValue(CValeur));
              if (index>-1) then
              begin
                 CBValeur.value := GetValue(CValeur);
                 CBValeur.text := CBValeur.Items[index];
              end else
                 CBValeur.text := GetValue(CValeur);
            end else CBValeur.text := GetValue(CValeur);
         end else
         if ((Copy(C.Name,5,2)='OP') or (Copy(C.Name,5,2)='LI')) then
         begin
            THValcomboBox(GetControl(C.Name)).value := GetValue(C.Name);
         end;
      end;
    end;
end;

function TOF_RTConfident.GetValue(Champtable: String): string;
begin
    while (Champtable[Length(Champtable)]='_') do
      System.Delete(Champtable,Length(Champtable),1);
    Result := TobAcces.GetValue(Champtable);
end;

Procedure TOF_RTConfident.PutValue(Champtable:string; Valeur: Variant);
begin
    while (Champtable[Length(Champtable)]='_') do
      System.Delete(Champtable,Length(Champtable),1);
    TobAcces.PutValue(Champtable,Valeur);
end;

procedure TOF_RTConfident.ControleTousCritere;
begin
StWhereCons  := ControleCritere('CON',TobConsult);
StWhereCons2 := ControleCritere('CO2',TobConsult2);
StWhereMod   := ControleCritere('MOD',TobModif);
StWhereMod2  := ControleCritere('MO2',TobModif2);
stWhereAct   := ControleCritere('ACT',TobAction);
stWherePro   := ControleCritere('CPR',TobCreatPro);
stWhereAction:= ControleCritere('CAC',TobCreatAction);
stWhereCon   := ControleCritere('CCO',TobCreatCon);
stWhereCic   := ControleCritere('CIC',TobCreatInfos);
end;

function TOF_RTConfident.ControleCritere (TypeConf: String;TobConf:TOB):string;
var     Champ,Oper,Val,Lien : Array[1..10] of integer ;
        C : TControl ;
        P : TTabSheet;
        i,n_ligne : integer;
        COper :string;
begin
  { Pour l'utilisation de RecupWhereCritere, on doit copier les composants de
    Panel typeConf :'PCON'    dans          PageControl 'PAGES'  avec son numéro de controle
    RTC_CHAMP1: THValComboBox        Z_C1                      Champ [10]
    RTC_OPER1: THValComboBox         Z01                       Oper  [10]
    RTC_VAL1: THValComboBox          ZV1                       Val   [10]
    RTC_LIEN1: THValComboBox         ZG1                       Lien  [10]}

    TobAcces.InitValeurs(false);
    TobAcces.putvalue('RTC_TYPECONF',TypeConf);
    if (TypeConf = 'CPR') then
    begin
    if TCheckBox(GetControl('CREATPROPO')).checked then TobAcces.putvalue('RTC_VAL1','=');
    end
    else
    if (TypeConf = 'CAC') then
    begin
    if TCheckBox(GetControl('CREATACTION')).checked then TobAcces.putvalue('RTC_VAL1','=');
    end
    else
    if (TypeConf = 'CCO') then
    begin
    if TCheckBox(GetControl('CREATCONTACT')).checked then TobAcces.putvalue('RTC_VAL1','=');
    end
    else
    if (TypeConf = 'CIC') then
    begin
    if TCheckBox(GetControl('CREATINFOS')).checked then TobAcces.putvalue('RTC_VAL1','=');
    end
    else
    if (TypeConf = 'ACT') then
    begin
      TobAcces.putvalue('RTC_CHAMP1','RTC_TYPEACTION');
      COper := THRadioGroup(GetControl('RGOPERATE')).value;
      TobAcces.putvalue('RTC_OPER1',COper);
      AutoriseAct := (COper = '=');
      TobAcces.putvalue('RTC_CHAMP2','RTC_UTILISATEUR');
      if TCheckBox(GetControl('CBCREATEUR')).checked then COper:= '='
      else COper := '<>';
      TobAcces.putvalue('RTC_OPER2',COper);
      if TCheckBox(GetControl('CBTOUSTYPES')).checked then TobAcces.putvalue('RTC_VAL1','=')
      else TobAcces.putvalue('RTC_VAL1','<>');
      ResponsableAct := (COper = '=');
      stWhere := '';
      for i:=0 to LBACTCODESEL.Items.Count-1 do
        stWhere := stWhere + LBACTCODESEL.items[i] + ';';
      TobAcces.putvalue('RTC_SQLCONF',stWhere);
      if (DroitProfil='AUCUN') and (stwhere='') and not ResponsableAct then Result := '-'
      else if (CBProfil.ItemIndex=0) and (DroitIntervenant='AUCUN') and (stwhere='') and not ResponsableAct
      then Result := '-'
      else Result := stwhere;
    end else
    begin
      P:= TPageControl(getcontrol('PAGES')).pages[0];
      for i:=0 to P.ControlCount-1 do
      begin
        C:=P.Controls[i] ;
        if Copy(C.Name,1,3)='Z_C' then Champ[StrToInt(Copy(C.Name,4,1))]:=i ;
        if Copy(C.Name,1,2)='ZO'  then Oper[StrToInt(Copy(C.Name,3,1))]:=i ;
        if Copy(C.Name,1,2)='ZV'  then Val[StrToInt(Copy(C.Name,3,1))]:=i ;
        if Copy(C.Name,1,2)='ZG'  then Lien[StrToInt(Copy(C.Name,3,1))]:=i ;
      end;

      PCons := TPanel(getcontrol('PTIERS'+TypeConf));
      for i:=1 to PCons.ControlCount-1 do
      begin
         C:=PCons.Controls[i] ;
         n_ligne := 1;
         if (TypeConf = 'CON') then
           n_ligne:=StrToint(Copy(C.Name,Length(C.Name),1))
         else if (TypeConf = 'CO2') then
           n_ligne:=StrToint(Copy(C.Name,Length(C.Name)-1,1))
         else if (TypeConf = 'MOD') then
           n_ligne:=StrToint(Copy(C.Name,Length(C.Name)-2,1))
         else if (TypeConf = 'MO2') then
           n_ligne:=StrToint(Copy(C.Name,Length(C.Name)-3,1));

         if Copy(C.Name,5,2)='CH' then
         begin
           THValComboBox(P.Controls[Champ[n_ligne]]).ItemIndex := THValComboBox(C).ItemIndex;
           Putvalue(C.Name,THValComboBox(C).value);
         end else
         if Copy(C.Name,5,2)='OP' then
         begin
           THValComboBox(P.Controls[Oper[n_ligne]]).ItemIndex := THValComboBox(C).ItemIndex;
           Putvalue(C.Name,THValComboBox(C).value);
         end else
         if Copy(C.Name,5,2)='VA' then
         begin
           if (THValComboBox(C).ItemIndex = -1) then
           begin
             THEdit(P.Controls[Val[n_ligne]]).Text := THValComboBox(C).text;
             Putvalue(C.Name,THValComboBox(C).text);
           end else
           begin
             THEdit(P.Controls[Val[n_ligne]]).Text := THValComboBox(C).value;
             Putvalue(C.Name,THValComboBox(C).value);
           end;
         end else
         if Copy(C.Name,5,2)='LI' then
         begin
           THValComboBox(P.Controls[Lien[n_ligne]]).ItemIndex := THValComboBox(C).ItemIndex;
           Putvalue(C.Name,THValComboBox(C).text);
         end;
      end;
      stWhere := RecupWhereCritere (TPageControl(getcontrol('PAGES')));
      stWhere:=VireLeWhere(stWhere);
      if (length(stWhere)>1) and ((TypeConf='CO2') or (TypeConf='MO2')) then
        stWhere := ' OR' + copy(stWhere,5,length(stWhere)-4);
      TobAcces.putvalue('RTC_SQLCONF',stWhere);
      if (DroitProfil='AUCUN') and (stwhere='') then Result := '-'
      else if (CBProfil.ItemIndex=0) and (DroitIntervenant='AUCUN') and (stwhere='') {and not ResponsableAct}
      then Result := '-'
      else Result := stwhere;
    end;
   TobConf.Dupliquer(TobAcces,False,True,True);
end;

procedure TOF_RTConfident.OnArgument(Arguments: string);
begin
  inherited;
  { mng 18/10/07 FQ 10741 }
  if (V_PGI.MenuCourant = 93 ) then
    begin
    stProduitpgi:='GRF';
    Ecran.Caption := TraduireMemoire('Restriction fournisseurs');
    UpdateCaption(Ecran);
    SetControlVisible('CREATPROPO',false);
    end
  else
    stProduitpgi:='GRC';    
  ModifLot := False;
  bTriLibelle:=V_PGI.ComboSortedOnLibelle;
  V_PGI.ComboSortedOnLibelle:=false;
  if (Arguments = 'MODIFLOT') then  ModifLot := true
  else Intervenant := Arguments;

  CBIntervenant := THValcomboBox(GetControl('RTC_INTERVENANT'));
  CBIntervenant.OnClick := BINTERVENANTClick;
  if ModifLot then CBIntervenant.ItemIndex := 0
  else CBIntervenant.Value := Intervenant;
  CBProfil := THValcomboBox(GetControl('RTC_PROFIL'));
  CBProfil.OnClick := BPROFILClick;
  CBProfil.ItemIndex := 0;
  TToolbarButton97(GetControl('BSQL')).OnClick := BSQLClick;
  TToolbarButton97(GetControl('BACCES')).OnClick := BTOUTACCESClick;
  CBCHAMP1 := THValcomboBox(GetControl('RTC_CHAMP1'));
  CBCHAMP1.OnClick := BCHAMP1Click;
  CBCHAMP2 := THValcomboBox(GetControl('RTC_CHAMP2'));
  CBCHAMP2.OnClick := BCHAMP2Click;
  CBCHAMP3 := THValcomboBox(GetControl('RTC_CHAMP3'));
  CBCHAMP3.OnClick := BCHAMP3Click;

  CBCHAMP1_ := THValcomboBox(GetControl('RTC_CHAMP1_'));
  CBCHAMP1_.OnClick := BCHAMP1_Click;
  CBCHAMP2_ := THValcomboBox(GetControl('RTC_CHAMP2_'));
  CBCHAMP2_.OnClick := BCHAMP2_Click;
  CBCHAMP3_ := THValcomboBox(GetControl('RTC_CHAMP3_'));
  CBCHAMP3_.OnClick := BCHAMP3_Click;

  CBCHAMP1__ := THValcomboBox(GetControl('RTC_CHAMP1__'));
  CBCHAMP1__.OnClick := BCHAMP1__Click;
  CBCHAMP2__ := THValcomboBox(GetControl('RTC_CHAMP2__'));
  CBCHAMP2__.OnClick := BCHAMP2__Click;
  CBCHAMP3__ := THValcomboBox(GetControl('RTC_CHAMP3__'));
  CBCHAMP3__.OnClick := BCHAMP3__Click;

  CBCHAMP1___ := THValcomboBox(GetControl('RTC_CHAMP1___'));
  CBCHAMP1___.OnClick := BCHAMP1___Click;
  CBCHAMP2___ := THValcomboBox(GetControl('RTC_CHAMP2___'));
  CBCHAMP2___.OnClick := BCHAMP2___Click;
  CBCHAMP3___ := THValcomboBox(GetControl('RTC_CHAMP3___'));
  CBCHAMP3___.OnClick := BCHAMP3___Click;

  LBACT :=  TListBox(GetControl('LBACT'));
  LBACTSEL :=  TListBox(GetControl('LBACTSEL'));
  LBACTCODESEL :=  TListBox(GetControl('LBACTCODESEL'));
  TToolbarButton97(GetControl('BAJOUTER')).OnClick := BAJOUTERClick;
  TToolbarButton97(GetControl('BENLEVER')).OnClick := BENLEVERClick;
  LBACT.OnDblClick:=BAJOUTERClick;
  LBACTSEL.OnDblClick:=BENLEVERClick;
  TToolbarButton97(GetControl('BCOPIE')).OnClick :=BCOPIEACCESClick;
  MemoCons := TMemo(getcontrol('WHERECONS'));
  MemoMod  := TMemo(getcontrol('WHEREMOD'));

  if Assigned(TCheckBox(GetControl('CREATACTION'))) then
    TCheckBox(GetControl('CREATACTION')).onClick:=OnCreatActionClick;
end;

procedure TOF_RTConfident.chargeComboChamps( Prefixe,ListeDhControle : string;ListeChamp,ListeLibelle: TStrings);
var iTable,iChamp,i :integer;
    Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
    FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  table := Mcd.getTable(mcd.PrefixetoTable('T'));
  FieldList := Table.Fields;
  FieldList.Reset();
  While FieldList.MoveNext do
  begin
    if (Pos(ListeDhControle,(FieldList.Current as IFieldCOM).control)>0) and ((FieldList.Current as IFieldCOM).libelle<>'') then
      if ((FieldList.Current as IFieldCOM).libelle[1]<>'.') then
        begin
        ListeChamp.add((FieldList.Current as IFieldCOM).name);
        ListeLibelle.add ((FieldList.Current as IFieldCOM).libelle);
        end;
  end;
end;

procedure TOF_RTConfident.OnLoad;
var ListeChamp,ListeLibelle: TStrings;
begin
  inherited;
  TobAcces:=TOB.create ('PROSPECTCONF',NIL,-1);
  TobConsult:=TOB.create ('PROSPECTCONF',NIL,-1);
  TobConsult.putvalue ('RTC_TYPECONF','CON');
  TobConsult2:=TOB.create ('PROSPECTCONF',NIL,-1);
  TobConsult2.putvalue ('RTC_TYPECONF','CO2');

  TobModif:=TOB.create ('PROSPECTCONF',NIL,-1);
  TobModif.putvalue ('RTC_TYPECONF','MOD');
  TobModif2:=TOB.create ('PROSPECTCONF',NIL,-1);
  TobModif2.putvalue ('RTC_TYPECONF','MO2');

  TobCreatPro:=TOB.create ('PROSPECTCONF',NIL,-1);
  TobCreatPro.putvalue ('RTC_TYPECONF','CPR');
  TobCreatCon:=TOB.create ('PROSPECTCONF',NIL,-1);
  TobCreatCon.putvalue ('RTC_TYPECONF','CCO');
  TobCreatInfos:=TOB.create ('PROSPECTCONF',NIL,-1);
  TobCreatInfos.putvalue ('RTC_TYPECONF','CIC');
  TobCreatAction:=TOB.create ('PROSPECTCONF',NIL,-1);
  TobCreatAction.putvalue ('RTC_TYPECONF','CAC');

  TobAction:=TOB.create ('PROSPECTCONF',NIL,-1);
  TobAction.putvalue ('RTC_TYPECONF','ACT');

  ListeChamp := TStringList.Create;
  ListeLibelle := TStringList.Create;
  LACTCODE := TStringList.Create;
  ListeChamp.add (''); ListeLibelle.add (TraduireMemoire('<<Aucun>>'));
  if stProduitpgi='GRC' then
    begin
    chargeComboChamps ('T','9',ListeChamp,ListeLibelle);
    chargeComboChamps ('YTC','9',ListeChamp,ListeLibelle);
    chargeComboChamps ('RPR','9',ListeChamp,ListeLibelle);
    end
  else
    begin
    chargeComboChamps ('T','8',ListeChamp,ListeLibelle);
    chargeComboChamps ('YTC','8',ListeChamp,ListeLibelle);
    chargeComboChamps ('RD3','8',ListeChamp,ListeLibelle);
    end ;

  CBCHAMP1.Items.Assign(ListeLibelle);
  CBCHAMP1.Values.Assign(ListeChamp);
  CBCHAMP2.Items.Assign(ListeLibelle);
  CBCHAMP2.Values.Assign(ListeChamp);
  CBCHAMP3.Items.Assign(ListeLibelle);
  CBCHAMP3.Values.Assign(ListeChamp);

  CBCHAMP1_.Items.Assign(ListeLibelle);
  CBCHAMP1_.Values.Assign(ListeChamp);
  CBCHAMP2_.Items.Assign(ListeLibelle);
  CBCHAMP2_.Values.Assign(ListeChamp);
  CBCHAMP3_.Items.Assign(ListeLibelle);
  CBCHAMP3_.Values.Assign(ListeChamp);

  CBCHAMP1__.Items.Assign(ListeLibelle);
  CBCHAMP1__.Values.Assign(ListeChamp);
  CBCHAMP2__.Items.Assign(ListeLibelle);
  CBCHAMP2__.Values.Assign(ListeChamp);
  CBCHAMP3__.Items.Assign(ListeLibelle);
  CBCHAMP3__.Values.Assign(ListeChamp);

  CBCHAMP1___.Items.Assign(ListeLibelle);
  CBCHAMP1___.Values.Assign(ListeChamp);
  CBCHAMP2___.Items.Assign(ListeLibelle);
  CBCHAMP2___.Values.Assign(ListeChamp);
  CBCHAMP3___.Items.Assign(ListeLibelle);
  CBCHAMP3___.Values.Assign(ListeChamp);

  THValcomboBox(GetControl('Z_C1')).Items.Assign(ListeLibelle);
  THValcomboBox(GetControl('Z_C1')).Values.Assign(ListeChamp);
  THValcomboBox(GetControl('Z_C2')).Items.Assign(ListeLibelle);
  THValcomboBox(GetControl('Z_C2')).Values.Assign(ListeChamp);
  THValcomboBox(GetControl('Z_C3')).Items.Assign(ListeLibelle);
  THValcomboBox(GetControl('Z_C3')).Values.Assign(ListeChamp);
  if (stProduitpgi = 'GRC' ) then
      RemplirValCombo('RTTYPEACTION','','',ListeLibelle,LACTCODE,False,False)
  else
      RemplirValCombo('RTTYPEACTIONFOU','','',ListeLibelle,LACTCODE,False,False);
  LBACT.Items.Assign(ListeLibelle);

  ListeChamp.free;ListeLibelle.Free;
  ChargeIntervenant;
  SignaleTTDroitModif(False);
  if Assigned(TCheckBox(GetControl('CREATACTION'))) then
    GriseModAction (TCheckBox(GetControl('CREATACTION')).checked);
end;

procedure TOF_RTConfident.OnCreatActionClick (Sender: TObject);
begin
  GriseModAction (TCheckBox(GetControl('CREATACTION')).checked);
end;

procedure TOF_RTConfident.GriseModAction (coche : boolean);
begin
  { si pas de création d'action, l'option modification limitée à ses propres
    actions ne sert à rien }
  if coche and (assigned(TCheckBox(GetControl('CBCREATEUR')))) then
    TCheckBox(GetControl('CBCREATEUR')).enabled:=true
  else
    begin
    TCheckBox(GetControl('CBCREATEUR')).enabled:=false;
    TCheckBox(GetControl('CBCREATEUR')).checked:=false;
    end;
end;

procedure TOF_RTConfident.OnClose;
begin
  inherited;
   TobAcces.free; TobConsult.free; TobModif.free; TobAction.free;
   TobConsult2.free; TobModif2.free;
   TobCreatPro.free; TobCreatCon.free; TobCreatInfos.free; TobCreatAction.free;
   LACTCODE.free;
  V_PGI.ComboSortedOnLibelle:=bTriLibelle;
end;

procedure TOF_RTConfident.OnUpdate;
begin
  inherited;
  Intervenant := CBIntervenant.value;
  if not ModifLot and (Intervenant='') then exit;
  ControleTousCritere;
  if ((stwhereCons='-') and (stwhereMod='-') {and (stwhereAct='-') and not ResponsableAct}) then ModifieLesAcces
  else if ((stwhereCons='') and (stwhereMod='') and (stwhereAct='') and not ResponsableAct) then DonneToutacces
  else begin
    if (Length(stWhereConsAvant)<=1) and (length(stWhereCons2Avant)<=1) and (Length(stwhereMod)<=1) and (length(stwhereMod2)<=1) then CopieAccesCons
    else if ( ((Length(stwhereCons)>1) and (stwhereCons<>stWhereConsAvant) and (stwhereCons<>stwhereMod))
    or   ((Length(stwhereCons2)>1) and (stwhereCons2<>stWhereCons2Avant) and (stwhereCons2<>stwhereMod2)) )
    and (HShowMessage(FMess[2], Ecran.Caption, '') <> mrYes) then exit;

    EnregistreIntervenant(0,True);
    if ModifLot then TFVierge(ecran).close
    else begin
      ChargeIntervenant;
      SignaleTTDroitModif(False);
    end;
  end;
end;

procedure TOF_RTConfident.OnDelete;
begin
  inherited;
  Intervenant := CBIntervenant.value;
  if not ModifLot and (Intervenant='') then exit;
  if SupprimeToutacces then
  begin
  CBProfil.ItemIndex := 0;DroitProfil:='';
  THLABEL(Getcontrol('TTDROITPROFIL')).caption := '';
  end;
  TFVierge(ecran).BFermeClick(nil);
end;

function TOF_RTConfident.SupprimeToutacces:Boolean;
begin
  result := False;
  if (HShowMessage(FMess[0], Ecran.Caption, '') = mrYes) then
    begin
    //EnregistreIntervenant(0,False);
    deleteIntervenant;
    if stProduitpgi = 'GRC' then
      begin
      VH_RT.DroitModifTiers:=false;
      VH_RT.RTConfWhereConsult:='';
      VH_RT.RTConfWhereModif:='';
      VH_RT.RTCreatPropositions := False;
      VH_RT.RTCreatActions := False;
      VH_RT.RTCreatContacts := False;
      VH_RT.RTCreatInfos := False;
      VH_RT.RTExisteConfident := False;
      end
    else
      begin
      VH_RT.RFConfWhereConsult:='';
      VH_RT.RFConfWhereModif:='' ;
      VH_RT.RFDroitModifTiers := False;
      VH_RT.RFExisteConfident := False;
      VH_RT.RFCreatContacts := False;
      VH_RT.RFCreatInfos := False;
      VH_RT.RFCreatActions := False;
      end;
    if ModifLot then TFVierge(ecran).close
    else begin
      ChargeIntervenant;
      SignaleTTDroitModif(False);
      Result := True;
    end;
  end;
end;

function TOF_RTConfident.ModifieLesAcces:Boolean;
begin
  result := False;
  if (HShowMessage(FMess[3], Ecran.Caption, '') = mrYes) then
  begin
    EnregistreIntervenant(0,False);
    //deleteIntervenant;
    if ModifLot then TFVierge(ecran).close
    else begin
      ChargeIntervenant;
      SignaleTTDroitModif(False);
      Result := True;
    end;
  end;
end;


procedure TOF_RTConfident.deleteIntervenant;
var stMess : string;
begin
  if (Intervenant = '') or (stProduitpgi = '') then
    begin
    PgiInfo('Suppression des droits d''accès impossible (code intervenant vide)','Droits de consultation uniquement');
    exit;
    end;

  if ExisteSql('SELECT US_UTILISATEUR FROM UTILISAT WHERE US_UTILISATEUR = "'+Intervenant+'"') then
    if ExecuteSql ('DELETE FROM PROSPECTCONF WHERE RTC_INTERVENANT = "'+Intervenant+'" AND RTC_PRODUITPGI = "'+stProduitpgi+'"') <> 0 then
      begin
      stMess:='Suppression des droits d''accès effectuée sur '+Intervenant;
      PgiInfo(stMess,'Droits de consultation uniquement');
      end;
end;

function TOF_RTConfident.DonneToutacces:Boolean;
begin
  result := False;
  if (HShowMessage(FMess[1], Ecran.Caption, '') = mrYes)  then
  begin
    TobConsult.InitValeurs(false); TobConsult2.InitValeurs(false);
    TobModif.InitValeurs(false);   TobModif2.InitValeurs(false);
    TobCreatPro.InitValeurs(false); TobCreatCon.InitValeurs(false); TobCreatInfos.InitValeurs(false);
    TobAction.InitValeurs(false); TobCreatAction.InitValeurs(false);
    TobConsult.putvalue  ('RTC_TYPECONF','CON');
    TobConsult2.putvalue ('RTC_TYPECONF','CO2');
    TobModif.putvalue    ('RTC_TYPECONF','MOD');
    TobModif2.putvalue   ('RTC_TYPECONF','MO2');
    TobAction.putvalue   ('RTC_TYPECONF','ACT');
    TobCreatPro.putvalue   ('RTC_TYPECONF','CPR');
    TobCreatCon.putvalue   ('RTC_TYPECONF','CCO');
    TobCreatInfos.putvalue   ('RTC_TYPECONF','CIC');
    TobCreatAction.putvalue   ('RTC_TYPECONF','CAC');
    if TCheckBox(GetControl('CBTOUSTYPES')).checked then TobAction.putvalue('RTC_VAL1','=')
    else TobAction.putvalue('RTC_VAL1','<>');
    if TCheckBox(GetControl('CREATPROPO')).checked then TobCreatPro.putvalue('RTC_VAL1','=')
    else TobAction.putvalue('RTC_VAL1','<>');
    if TCheckBox(GetControl('CREATACTION')).checked then TobCreatAction.putvalue('RTC_VAL1','=')
    else TobCreatAction.putvalue('RTC_VAL1','<>');

    if TCheckBox(GetControl('CREATCONTACT')).checked then TobCreatCon.putvalue('RTC_VAL1','=')
    else TobAction.putvalue('RTC_VAL1','<>');
    if TCheckBox(GetControl('CREATINFOS')).checked then TobCreatInfos.putvalue('RTC_VAL1','=')
    else TobAction.putvalue('RTC_VAL1','<>');

    EnregistreIntervenant(0,True);
    if ModifLot then TFVierge(ecran).close
    else begin
      ChargeIntervenant;
      SignaleTTDroitModif(False);
      Result := True;
    end;
  end;
end;

procedure TOF_RTConfident.BCOPIEACCESClick(Sender: TObject);
begin
  ControleTousCritere;
  CopieAccesCons;
end;

procedure TOF_RTConfident.CopieAccesCons;
begin
  TobModif.Dupliquer (TobConsult,False,True,True);
  TobModif.putvalue  ('RTC_TYPECONF','MOD');
  TobAcces.InitValeurs(false);
  TobAcces.Dupliquer (TobModif,False,True,True);
  AfficheCritere('MOD');
  TobModif2.Dupliquer(TobConsult2,False,True,True);
  TobModif2.putvalue ('RTC_TYPECONF','MO2');
  TobAcces.InitValeurs(false);
  TobAcces.Dupliquer (TobModif2,False,True,True);
  AfficheCritere('MO2');
end;


procedure TOF_RTConfident.BTOUTACCESClick(Sender: TObject);
begin
  Intervenant := CBIntervenant.value;
  if not ModifLot and (Intervenant='') then exit;
  if DonneToutacces then
  begin
  CBProfil.ItemIndex := 0;DroitProfil:='';
  THLABEL(Getcontrol('TTDROITPROFIL')).caption := '';
  end;
end;

procedure TOF_RTConfident.SignaleTTDroitModif(profil :boolean);
var TTDroit : THLABEL ; 
    AucunDroit,ToutDroit : boolean;
begin
 AucunDroit := ((stwhereCons='-') and (stwhereMod='-') {and (stwhereAct='-') and not (ResponsableAct)});
 ToutDroit :=  ((stwhereCons='' ) and (stwhereMod='' ) {and (stwhereAct='' ) and not (ResponsableAct)});
 if profil then
 begin
   TTDroit := THLABEL(Getcontrol('TTDROITPROFIL')); TTDroit.font.Style := [fsBold];
   TTDroit.caption := '';  DroitProfil:='';
   if AucunDroit then
   begin     //droits de consultation uniquement
      TTDroit.caption := TraduireMemoire('Droits de consultation uniquement');DroitProfil:='AUCUN'; TTDroit.font.color:=clRed;
   end;
   if ToutDroit then
   begin
      TTDroit.caption := TraduireMemoire('Tous les droits');DroitProfil:='TOUS'; TTDroit.font.color:=clGreen;
   end
 end else
 begin
   TTDroit := THLABEL(Getcontrol('TTDROITUTIL')); TTDroit.font.Style := [fsBold];
   TTDroit.caption := ''; DroitIntervenant := '';
   if (Intervenant<>'') and AucunDroit then
   begin
      DroitIntervenant := 'AUCUN';
      TTDroit.caption := TraduireMemoire('Droits de consultation uniquement');TTDroit.font.color:=clRed;
   end;
   if (Intervenant<>'') and ToutDroit then
   begin
      DroitIntervenant := 'TOUS';
      TTDroit.caption := TraduireMemoire('Tous les droits');TTDroit.font.color:=clGreen;
   end;
 end;
end;

procedure TOF_RTConfident.EnregistreIntervenant(count : integer;Maj:boolean);
var Consult1,Consult2,Modif1,Modif2 :string;
     stConfWhereConsult : string ;
     stConfWhereModif : string ;
     bDroitModifTiers : boolean ;
begin
  if ModifLot then Intervenant := ListeIntervenant[count];
  TobConsult.putvalue('RTC_INTERVENANT',Intervenant);
  TobConsult.putvalue('RTC_PRODUITPGI',stProduitpgi);
  TobConsult.SetAllModifie (True);
  TobConsult.ChargeCle1;
  if TobConsult.ExistDB then TobConsult.DeleteDB;
  if Maj then
  TobConsult.InsertOrUpdateDB;

  TobConsult2.putvalue('RTC_INTERVENANT',Intervenant);
  TobConsult2.putvalue('RTC_PRODUITPGI',stProduitpgi);
  TobConsult2.SetAllModifie (True);
  TobConsult2.ChargeCle1;
  if TobConsult2.ExistDB then TobConsult2.DeleteDB;
  if Maj then
  TobConsult2.InsertOrUpdateDB;

  TobModif.putvalue('RTC_INTERVENANT',Intervenant);
  TobModif.putvalue('RTC_PRODUITPGI',stProduitpgi);
  TobModif.SetAllModifie (True);
  TobModif.ChargeCle1;
  if TobModif.ExistDB then TobModif.DeleteDB;
  if Maj then
  TobModif.InsertOrUpdateDB;

  TobModif2.putvalue('RTC_INTERVENANT',Intervenant);
  TobModif2.putvalue('RTC_PRODUITPGI',stProduitpgi);
  TobModif2.SetAllModifie (True);
  TobModif2.ChargeCle1;
  if TobModif2.ExistDB then TobModif2.DeleteDB;
  if Maj then
  TobModif2.InsertOrUpdateDB;

  TobCreatPro.putvalue('RTC_INTERVENANT',Intervenant);
  TobCreatPro.putvalue('RTC_PRODUITPGI',stProduitpgi);
  TobCreatPro.SetAllModifie (True);
  TobCreatPro.ChargeCle1;
  if TobCreatPro.ExistDB then TobCreatPro.DeleteDB;
  //if Maj then
  TobCreatPro.InsertOrUpdateDB;

  TobCreatAction.putvalue('RTC_INTERVENANT',Intervenant);
  TobCreatAction.putvalue('RTC_PRODUITPGI',stProduitpgi);
  TobCreatAction.SetAllModifie (True);
  TobCreatAction.ChargeCle1;
  if TobCreatAction.ExistDB then TobCreatAction.DeleteDB;
  //if Maj then
  TobCreatAction.InsertOrUpdateDB;

  TobCreatCon.putvalue('RTC_INTERVENANT',Intervenant);
  TobCreatCon.putvalue('RTC_PRODUITPGI',stProduitpgi);
  TobCreatCon.SetAllModifie (True);
  TobCreatCon.ChargeCle1;
  if TobCreatCon.ExistDB then TobCreatCon.DeleteDB;
  //if Maj then
  TobCreatCon.InsertOrUpdateDB;

  TobCreatInfos.putvalue('RTC_INTERVENANT',Intervenant);
  TobCreatInfos.putvalue('RTC_PRODUITPGI',stProduitpgi);
  TobCreatInfos.SetAllModifie (True);
  TobCreatInfos.ChargeCle1;
  if TobCreatInfos.ExistDB then TobCreatInfos.DeleteDB;
  //if Maj then
  TobCreatInfos.InsertOrUpdateDB;

  TobAction.putvalue('RTC_INTERVENANT',Intervenant);
  TobAction.putvalue('RTC_PRODUITPGI',stProduitpgi);
  TobAction.SetAllModifie (True);
  TobAction.ChargeCle1;
  if TobAction.ExistDB then TobAction.DeleteDB;
  //if Maj then
  TobAction.InsertOrUpdateDB;
  bDroitModifTiers := maj;

  if  (Intervenant = V_PGI.User ) then
  begin
    Consult1 := TobConsult.GetValue ('RTC_SQLCONF');
    Consult2 := TobConsult2.GetValue('RTC_SQLCONF');
    Modif1   := TobModif.GetValue   ('RTC_SQLCONF');
    Modif2   := TobModif2.GetValue  ('RTC_SQLCONF');
    if Consult2 = '' then  stConfWhereConsult := Consult1
   //mng 25/06/03 else VH_RT.RTConfWhereConsult := copy(Consult1,5,(Length(Consult1)-4)) + Consult2 ;
   else stConfWhereConsult := ' AND ('+copy(Consult1,5,(Length(Consult1)-4)) + Consult2 +')' ;

    if Modif2 = '' then  stConfWhereModif := Modif1
    else stConfWhereModif := ' AND ('+copy(Modif1,5,(Length(Modif1)-4)) + Modif2 +')';
  end;
  if stProduitpgi = 'GRC' then
    begin
    VH_RT.DroitModifTiers:=bDroitModifTiers;
    VH_RT.RTConfWhereConsult:=stConfWhereConsult;
    VH_RT.RTConfWhereModif:=stConfWhereModif;
    VH_RT.RTCreatPropositions:=(TobCreatPro.GetValue ('RTC_VAL1') = '=');
    VH_RT.RTCreatActions:=(TobCreatAction.GetValue ('RTC_VAL1') = '=');
    VH_RT.RTCreatContacts:=(TobCreatCon.GetValue ('RTC_VAL1') = '=');
    VH_RT.RTCreatInfos:=(TobCreatInfos.GetValue ('RTC_VAL1') = '=');
    end
  else
    begin
    VH_RT.RFDroitModifTiers:=bDroitModifTiers;
    VH_RT.RFConfWhereConsult:=stConfWhereConsult;
    VH_RT.RFConfWhereModif:=stConfWhereModif;
    VH_RT.RFCreatActions:=(TobCreatAction.GetValue ('RTC_VAL1') = '=');
    VH_RT.RFCreatContacts:=(TobCreatCon.GetValue ('RTC_VAL1') = '=');
    VH_RT.RFCreatInfos:=(TobCreatInfos.GetValue ('RTC_VAL1') = '=');
    end;

  if (modiflot) and  (count < ListeIntervenant.Count-1) then
  begin
    inc(count);
    EnregistreIntervenant(count,Maj);
  end;
end;

procedure TOF_RTMULConfident.OnArgument (S : String ) ;
begin
  if (V_PGI.MenuCourant = 93 ) then
    begin
    Ecran.Caption := TraduireMemoire('Restriction fournisseurs');
    UpdateCaption(Ecran);
    end;
end;

procedure TOF_RTMULConfident.ChargeMulIntervenant(Intervenant:string);
Var F : TFMul ;
    i : integer;
begin
  F:=TFMul(Ecran);
  if(F.FListe.NbSelected<2)and(not F.FListe.AllSelected) then
    AGLLanceFiche('RT','RTCONFIDENT','','',Intervenant)
  else
  begin
    ListeIntervenant := TStringList.Create;
    if F.FListe.AllSelected then
    begin
      F.Q.DisableControls ;
      F.Q.First;
      while Not F.Q.EOF do
      begin
        Intervenant := F.Q.FindField('US_UTILISATEUR').AsString;
        with ListeIntervenant do Add (Intervenant);
        F.Q.Next;
      end;
      AGLLanceFiche('RT','RTCONFIDENT','','','MODIFLOT') ;
      F.FListe.AllSelected:=False
    end else
    begin
      for i:=0 to F.FListe.NbSelected-1 do
      begin
        F.FListe.GotoLeBookMark(i) ;
{$IFDEF EAGLCLIENT}
        F.Q.TQ.Seek(F.FListe.Row-1) ;
{$ENDIF}
        Intervenant := F.Q.FindField('US_UTILISATEUR').AsString;
        with ListeIntervenant do Add (Intervenant);
      end;
      AGLLanceFiche('RT','RTCONFIDENT','','','MODIFLOT') ;
      F.FListe.ClearSelected;
    end;
    F.Q.EnableControls ;
    ListeIntervenant.free;
  end;
end;

/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLConfidentProspect(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
     Intervenant :string;
begin
F:=TForm(Longint(Parms[0])) ;
Intervenant := Parms[1];
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_RTMULConfident) then TOF_RTMULConfident(TOTOF).ChargeMulIntervenant(Intervenant) else exit;
end;

Initialization
registerclasses([Tof_RTConfident]);
registerclasses([Tof_RTMULConfident]);
RegisterAglProc('ConfidentProspect',TRUE,2,AGLConfidentProspect);
end.
