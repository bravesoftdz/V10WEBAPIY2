unit UTofMBOTransImpr;

interface
uses  M3FP,StdCtrls,Controls,Classes,HTB97,
{$IFDEF EAGLCLIENT}
      eQRS1,Maineagl,
{$ELSE}
      dbTables,DBGrids,db,QRS1,FE_Main,
{$ENDIF}
{$IFDEF AGL_SANS_RELOADCOMBOETAT}
      Windows, // Inclus temporairement pour fct ReloadComboEtat
{$ENDIF}
      forms,sysutils,ComCtrls,
      HDB,HCtrls,HEnt1,HMsgBox,UTOF, vierge, HDimension,UTOB, UtilArticle,UDimArticle,
      AglInit,EntGC,utilPGI;



Type
     TOF_MBOTransImpr = Class (TOF)
     public
        procedure OnArgument (Arguments : String ) ; override ;
        procedure OnUpdate ; override ;
        procedure OnClose ; override ;
     private
        TobTypeMasque : TOB ; // Tob des types de masques
        bTypeMasqueVisible,ChoixEdition : boolean ;
        procedure AjusteControle(bEnabled : boolean) ;
{$IFDEF AGL_SANS_RELOADCOMBOETAT}
        procedure ReloadComboEtat(lequel : string = '') ; //Tempo : attente AGL 550e
{$ENDIF}
        procedure DefinirModeleEtat ;
        procedure ReloadTypeMasque(TypeEtat : String) ;
        procedure MajMultiEtab(Combo : String) ;
        procedure ValiderChoix;
        procedure FermerChoix;
     END ;

implementation

procedure TOF_MBOTransImpr.OnArgument(Arguments : String ) ;
var QQ : TQuery ;
    Critere,ChampMul,ValMul,TypeDeLEtat,Choix : string;
    x : integer;
begin
inherited ;
if TForm(Ecran).name<>'CHOIXEDITTRANS' then
   begin
   // Chargement en Tob des types de masques
   QQ:=OpenSQL('select GMQ_TYPEMASQUE,GMQ_MULTIETAB from TYPEMASQUE',True) ;
   TobTypeMasque:=TOB.Create('TYPEMASQUE',nil,-1) ;
   TobTypeMasque.LoadDetailDB('TYPEMASQUE','','',QQ,False) ;
   Ferme(QQ) ;

   // Initialisation du choix de l'édition
   Repeat
   Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
   if Critere<>'' then
      begin
      x:=pos('=',Critere);
      if x<>0 then
         begin
         ChampMul:=copy(Critere,1,x-1);
         ValMul:=copy(Critere,x+1,length(Critere));
         if (ChampMul='RBBONPREPARATION') or (ChampMul='RBPROPOSITION') then
            begin
            SetControlChecked(ChampMul,StringToCheck(ValMul));
            if StringToCheck(ValMul) then TypeDeLEtat:=ChampMul;
            end;
         end;
      end;
   until  Critere='';
   ReloadTypeMasque(TypeDeLEtat);
   if TCheckBox(GetControl('RBBONPREPARATION')).checked then
   // Affichage ou non du critère Type de masque
   bTypeMasqueVisible:=((TobTypeMasque<>nil) and (TobTypeMasque.Detail.Count>1))
   else bTypeMasqueVisible:=False;
   SetControlVisible('TYPEMASQUE',bTypeMasqueVisible) ;
   SetControlVisible('LTYPEMASQUE',bTypeMasqueVisible) ;
   SetControlText('TYPEMASQUE',VH_GC.BOTypeMasque_Defaut) ;
   end
else
    begin
    // Lecture dans la registry du dernier choix effectué 
    Choix:=GetSynRegKey('CHOIXEDT', '', True) ;
    if Choix='BONPREPA' then SetControlChecked('RBBONPREPARATION',StringToCheck('X'))
    else if Choix='PROPO' then SetControlChecked('RBPROPOSITION',StringToCheck('X'));
    end;
end ;

procedure TOF_MBOTransImpr.OnUpdate ;
begin
inherited ;
if TForm(Ecran).name<>'CHOIXEDITTRANS' then
   QRDIMTypeMask:=GetControlText('TYPEMASQUE');
end ;

procedure TOF_MBOTransImpr.OnClose ;
begin
inherited ;
if TForm(Ecran).name<>'CHOIXEDITTRANS' then
   begin
   QRDIMTypeMask:=VH_GC.BOTypeMasque_Defaut ; // Réinitialisation à la valeur par défaut
   if TobTypeMasque<>nil then BEGIN TobTypeMasque.free ; TobTypeMasque:=nil END ;
   end
else if choixEdition then LastError:=0 else LastError:=-1;
end ;

procedure TOF_MBOTransImpr.AjusteControle(bEnabled : boolean) ;
begin
if bTypeMasqueVisible then
   BEGIN
   SetControlEnabled('TYPEMASQUE',bEnabled) ;
   SetControlEnabled('LTYPEMASQUE',bEnabled) ;
   END ;
end ;

{$IFDEF AGL_SANS_RELOADCOMBOETAT}
// Tempo : attente AGL 550e
procedure TOF_MBOTransImpr.ReloadComboEtat(lequel : string = '');
Var Q : TQuery ;
    i,idef : integer;
    Dernier,StSQL : string;
begin
TFQRS1(Ecran).FEtat.Items.Clear ; TFQRS1(Ecran).FEtat.Values.Clear ; idef:=-1;
StSQL:='SELECT MO_TYPE, MO_NATURE, MO_CODE, MO_LIBELLE, MO_LANGUE, MO_DEFAUT FROM MODELES WHERE MO_TYPE="'+TFQRS1(Ecran).FTypeEtat+'" AND MO_NATURE="'+TFQRS1(Ecran).FNatEtat+'"' ;
{if Not FChoixEtat then }
StSQL:=StSQL+' AND MO_CODE="'+TFQRS1(Ecran).FCodeEtat+'"' ;
Q:=OpenSQL(StSQL,TRUE) ;
While Not Q.EOF do
  BEGIN
  i:=TFQRS1(Ecran).FEtat.Items.Add(Q.FindField('MO_LIBELLE').AsString) ;
  if Q.FindField('MO_DEFAUT').AsString='X' then Idef:=i ;
  TFQRS1(Ecran).FEtat.Values.Add(Q.FindField('MO_CODE').AsString) ;
  Q.Next
  END ;
Ferme(Q) ;
TFQRS1(Ecran).FEtat.ItemIndex:=-1;
if lequel='' then
  begin
  Dernier:=GetFromRegistry(HKEY_LOCAL_MACHINE,'Software\'+Apalatys+'\'+NomHalley+'\QRS1Def',GetLeNom(TFQRS1(Ecran).FNomFiltre),'') ;
  if Dernier<>'' then TFQRS1(Ecran).FEtat.ItemIndex:= TFQRS1(Ecran).FEtat.values.IndexOf(Dernier);
  if TFQRS1(Ecran).FEtat.ItemIndex=-1 then TFQRS1(Ecran).FEtat.ItemIndex:=idef;
  end else TFQRS1(Ecran).FEtat.ItemIndex:= TFQRS1(Ecran).FEtat.values.IndexOf(lequel);
TFQRS1(Ecran).FEtatClick(Nil) ;
end;
{$ENDIF}

procedure TOF_MBOTransImpr.DefinirModeleEtat ;
var nature,code : string ;
begin
if TCheckBox(GetControl('RBBONPREPARATION')).Checked then
   BEGIN
   if THEdit(GetControl('MULTIETAB')).Text<>'X' then
      begin
      nature:='PBO';
      code:='ART';
      end
   else
      begin
      nature:='PBM' ;
      code:='ARM';
      end;
   END
   else
      begin
      nature:='PTR';
      code:='BTQ';
      end;
with TFQRS1(Ecran) do
   BEGIN
   NatureEtat:=nature ;
   CodeEtat:=code;
   ReloadComboEtat;
   END ;
end ;

procedure TOF_MBOTransImpr.ReloadTypeMasque(TypeEtat : String) ;
var ComboPlus : string ;
begin
if TypeEtat<>'RBBONPREPARATION' then
   BEGIN
   ComboPlus:='and (GMQ_MULTIETAB<>"X")' ;
   THEdit(GetControl('MULTIETAB')).Text:='-' ;
   AjusteControle(False) ;
   END
else
   BEGIN
   ComboPlus:='' ;
   THEdit(GetControl('MULTIETAB')).Text:='-' ; // '-' pour type masque défaut
   AjusteControle(True) ;
   END ;
with THValComboBox(GetControl('TYPEMASQUE')) do
   BEGIN
   Plus:=ComboPlus ;
   Reload ;
   Value:=VH_GC.BOTypeMasque_Defaut ;
   END ;
DefinirModeleEtat ;
end ;

procedure TOF_MBOTransImpr.MajMultiEtab(Combo : String) ;
var TobTM : Tob ;
begin
if Combo<>'TYPMAS' then exit ;
TobTM:=TobTypeMasque.FindFirst(['GMQ_TYPEMASQUE'],[THValComboBox(GetControl('TYPEMASQUE')).Value],False) ;
if TobTM<>nil then THEdit(GetControl('MULTIETAB')).Text:=TobTM.GetValue('GMQ_MULTIETAB') ;
DefinirModeleEtat ;
end ;

procedure TOF_MBOTransImpr.ValiderChoix;
var rbBonPreparation, rbProposition : TCheckBox ;
begin
rbBonPreparation:=TCheckBox(GetControl('RBBONPREPARATION'));
rbProposition:=TCheckBox(GetControl('RBPROPOSITION'));
if (rbBonPreparation.Checked) then
// Enregistrement dans la registry l'édition choisie
SaveSynRegKey('CHOIXEDT', 'BONPREPA', True);
if (rbProposition.Checked) then
// Enregistrement dans la registry l'édition choisie
SaveSynRegKey('CHOIXEDT', 'PROPO', True);
if ((rbBonPreparation.Checked) or (rbProposition.Checked)) then
    begin
    ChoixEdition:=True;
    AGLLanceFiche('MBO','TRANSIMPR','GTL_CODEPTRF='+THEdit(GetControl('CODEPTRF')).Text,'','NATURE='+THEdit(GetControl('NATURE')).text+';CODEETAT='+THEdit(GetControl('CODEETAT')).text+';RBBONPREPARATION='+CheckToString(rbBonPreparation.checked)+';RBPROPOSITION='+CheckToString(rbProposition.checked)) ;
    end
else
    begin
    HShowmessage('1;Choix;Vous devez choisir une édition.;W;O;O;O;','','');
    end;
end;

procedure TOF_MBOTransImpr.FermerChoix;
begin
ChoixEdition:=True;
end;

// Retourne la nouvelle nature état en fonction des nouveaux paramètres
procedure AGLOnChangeField (Parms: array of variant; nb: integer) ;
var F : TForm ;
    TOTOF : TOF ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFQRS1) then TOTOF:=TFQRS1(F).LaTOF else exit ;
if (TOTOF is TOF_MBOTransImpr) then
    if Parms[1]='RB' then TOF_MBOTransImpr(TOTOF).ReloadTypeMasque(Parms[2])
    else if Parms[1]='CB' then TOF_MBOTransImpr(TOTOF).MajMultiEtab(Parms[2])
    else exit ;
end;

procedure AGLValiderChoix (Parms: array of variant; nb: integer) ;
var F : TForm;
    TOTOF : TOF ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then TOTOF:=TFVierge(F).LaTOF else exit ;
if (TOTOF is TOF_MBOTransImpr) then TOF_MBOTransImpr(TOTOF).ValiderChoix;
end;

// Procédure servant à fermer la fiche Choix de l'édition
procedure AGLFermerChoix (Parms: array of variant; nb: integer) ;
var F : TForm;
    TOTOF : TOF ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then TOTOF:=TFVierge(F).LaTOF else exit ;
if (TOTOF is TOF_MBOTransImpr) then TOF_MBOTransImpr(TOTOF).FermerChoix;
end;

Initialization
registerclasses([TOF_MBOTransImpr]) ;
RegisterAglProc('OnChangeField', True , 2, AGLOnChangeField) ;
RegisterAglProc('ValiderChoix', True , 2, AGLValiderChoix) ;
RegisterAglProc('FermerChoix', True , 2, AGLFermerChoix) ;
end.

