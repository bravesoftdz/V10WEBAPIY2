unit UtofPlanningAct;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
{$IFDEF EAGLCLIENT}

{$ELSE}
   dbTables, db,HDB,
{$ENDIF}
      HCtrls,HEnt1,HMsgBox,UTOF, vierge,AffaireUtil,UTob,Grids,EntGC,
      DicoAF,Saisutil,Hstatus,M3FP,HPlanning,AfPlanningConfig;
Type
     TOF_PlanningAct = Class (TOF)
     public
        Planning : THPlanning;
        sCodePlanning : string;
        procedure OnArgument (stArgument : String ) ; override ;
        procedure OnLoad ; override ;
        procedure OnClose ; override ;
     private
        TOBRES,TOBACT,TobEtats,TobEvts : TOB;
        LibCrit : THLabel;
        LabelCrit : THLabel;
        LibDate : THLabel;
        EditCritere : THEdit;
        EditDate : THEdit;

        procedure AlimPlanning(AcbInit:boolean);
        procedure ConfigureEcran(TobConfigPlanning:TOB);

    END ;

implementation

procedure TOF_PlanningAct.OnArgument (stArgument : String ) ;
var
    Crit, Champ, valeur  : String;
    x : integer;
begin
  inherited;
// Recup des critères
Crit:=(Trim(ReadTokenSt(stArgument)));
While (Crit <>'') do
    BEGIN
    if Crit<>'' then
        BEGIN
        X:=pos('=',Crit);
        if x<>0 then
           begin
           Champ:=copy(Crit,1,X-1);
           Valeur:=Copy (Crit,X+1,length(Crit)-X);
           end;
        if Champ = 'CODE' then sCodePlanning := Valeur;
        END;
    Crit:=(Trim(ReadTokenSt(stArgument)));
    END;

{Planning := THPlanning(Ecran.FindComponent('PlanningAct'));
Planning.OnActive     := Activation ;
Planning.OnModifyItem := Modification ;
Planning.OnDeleteItem := Suppression;
Planning.OnDblClick   := DoubleClick;
Planning.OnMoveItem   := Deplacement ;
Planning.OnCopyItem   := CopyItem ;
Planning.OnInitItem   := InitItem ;
Planning.OnCreateItem := Creation;
//  Planning.AddOptionPopup(1,'Changer l''état de la résa.');
Planning.OnOptionPopup := OnPopup;    }

EditCritere := THEdit(GetControl('ACT_RESSOURCE'));
EditDate := THEdit(GetControl('ACT_DATEACTIVITE'));
LabelCrit := THLabel(GetControl('TACT_RESSOURCE'));
LibCrit := THLabel(GetControl('TACT_RESSOURCE1'));
LibDate := THLabel(GetControl('TACT_DATEACTIVITE'));

AlimPlanning(true);
end;

procedure TOF_PlanningAct.OnLoad;
Begin
 Inherited;

End;

procedure TOF_PlanningAct.OnClose;

Begin
 Inherited;

     if (TOBACT<>nil) then
        begin
        TOBACT.Free;TOBACT:=nil;
        end;
     if (TOBRES<>nil) then
        begin
        TOBRES.Free;TOBRES:=nil;
        end;
     if (TobEtats<>nil) then
        begin
        TobEtats.Free;TobEtats:=nil;
        end;
     if (TobEvts<>nil) then
        begin
        TobEvts.Free;TobEvts:=nil;
        end;

End;

procedure TOF_PlanningAct.AlimPlanning(AcbInit:boolean);
var TobPlanningEnCours,TobModelePlanning:TOB;
    Indice:Integer;
    sCritere:string;
    DateCrit : TDateTime;
begin
    sCritere := '';
    DateCrit := Date;

     TobModelePlanning:=TOB.Create ('les modeles', Nil, -1);
     TobModelePlanning.LoadDetailDB ('HRPARAMPLANNING', '', '',Nil, True) ;

     TobPlanningEnCours:=TobModelePlanning.FindFirst(['HPP_PARAMPLANNING'],[sCodePlanning],True);
     indice:=TobPlanningEnCours.GetIndex;

     if AcbInit and (TobPlanningEnCours<>nil) then ConfigureEcran(TobPlanningEnCours);


     if (EditCritere<>nil) then
        if (EditCritere.visible=true) then
            sCritere := EditCritere.Text;

     if (EditDate<>nil) then  
        DateCrit:=strtodate(EditDate.text);

     // Chargement des paramètres du planning
     try

     if (TOBACT<>nil) then
        begin
        TOBACT.Free;TOBACT:=nil;
        end;
     if (TOBRES<>nil) then
        begin
        TOBRES.Free;TOBRES:=nil;
        end;
     if (TobEtats<>nil) then
        begin
        TobEtats.Free;TobEtats:=nil;
        end;
     if (TobEvts<>nil) then
        begin
        TobEvts.Free;TobEvts:=nil;
        end;

     ChargeParamPlanning(Planning, TobModelePlanning, indice, sCritere, DateCrit, TobEtats, TOBACT, TOBRes, TobEvts);
     finally
     // Destruction de la tob des modeles de planning
     If TobModelePlanning <> Nil then
        begin
        TobModelePlanning.Free;
        end;
     end;
end;


procedure TOF_PlanningAct.ConfigureEcran(TobConfigPlanning:TOB);
begin

if (TobConfigPlanning.getValue('HPP_MODEGESTION')='AFF') then
    begin
    LabelCrit.Caption := TraduitGA('&Affaire');
    EditCritere.DataType := 'AFLAFFAIRE';
    end
else
if (TobConfigPlanning.getValue('HPP_MODEGESTION')='RES') then
    begin
    LabelCrit.Caption := TraduitGA('&Ressource');
    EditCritere.DataType := 'AFLRESSOURCE';
    end
else
if (TobConfigPlanning.getValue('HPP_MODEGESTION')='GLA') or
    (TobConfigPlanning.getValue('HPP_MODEGESTION')='GLR') then
    begin
    LabelCrit.Visible := false;
    LibCrit.Visible := false;
    EditCritere.Visible := false;
    LibDate.left := LabelCrit.left;
    EditDate.Left := EditCritere.Left;
    end;

Ecran.Caption := Ecran.Caption + ' ' + TobConfigPlanning.getValue('HPP_LIBELLE');
updatecaption(Ecran);
end;

// Fonctions AGL
procedure AGLAlimPlanning(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then MaTOF:=TFVierge(F).LaTOF else exit;
if (MaTOF is TOF_PlanningAct) then TOF_PlanningAct(MaTOF).AlimPlanning(false) else exit;
end;

Initialization
registerclasses([TOF_PlanningAct]);
RegisterAglProc('AlimPlanning',TRUE,0,AGLAlimPlanning);
end.
