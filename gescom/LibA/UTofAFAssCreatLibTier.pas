unit UTofAFAssCreatLibTier;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls, 
{$IFDEF EAGLCLIENT}
        Emul,Etablette, Maineagl,
{$ELSE}
       db,dbTables,HDB,DBCtrls, tablette,mul, FE_Main,
{$ENDIF}
      HCtrls,Ent1,HMsgBox,UTOF, DicoAF,SaisUtil,EntGC,HTB97,
      Buttons, ExtCtrls, HEnt1, utob, M3FP;
Type
     TOF_AFASSCREATLIBTIER = Class (TOF)
         procedure OnArgument(stArgument : String ) ; override ;
         procedure OnUpdate ; override ;
         procedure bCritere1Click(Sender: TObject);
         procedure bCritere2Click(Sender: TObject);
         procedure bCritere3Click(Sender: TObject);
         procedure bCritere4Click(Sender: TObject);
         procedure bCritere5Click(Sender: TObject);
         procedure bCritere6Click(Sender: TObject);
         procedure bCritere7Click(Sender: TObject);
         procedure bCritere8Click(Sender: TObject);
         procedure bCritere9Click(Sender: TObject);
         procedure bCritere10Click(Sender: TObject);
     public
        bCritLib1:boolean;
        bCritLib2:boolean;
        bRess1:boolean;
        bRess2:boolean;
        bRess3:boolean;
     END ;

Function  AFLanceFiche_CreatLibelClt(Argument:string):variant;

implementation


procedure TOF_AFASSCREATLIBTIER.OnArgument(stArgument : String );
var
Crit:string;
Begin
Inherited;
bCritLib1:=false;
bCritLib2:=false;
bRess1:=false;
bRess2:=false;
bRess3:=false;

Crit:=(Trim(ReadTokenSt(stArgument)));
while (Crit<>'') do
    begin
    if (Crit='BUREAU') then bCritLib1:=true;
    if (Crit='AGENCE') then bCritLib2:=true;
    if (Crit='ASSOCIE') then bRess1:=true;
    if (Crit='CHEFGROUPE') then bRess2:=true;
    if (Crit='RESPONSABLE') then bRess3:=true;

    Crit:=(Trim(ReadTokenSt(stArgument)));
    end;

if bCritLib1 then
    begin
    THEdit(GetControl('LIBCRIT1')).Text := 'Bureau';
    THEdit(GetControl('LIBCRIT1')).Enabled:=false;
    TToolBarButton97(GetControl('SBCRIT1')).Enabled:=false;
    end
else
//    THEdit(GetControl('LIBCRIT1')).Text := RechDom('GCCATTIERSTABLE','AL1',False);
    THEdit(GetControl('LIBCRIT1')).Text := RechDom('GCZONELIBRE','CT1',False);

if bCritLib2 then
    begin
    if bCritLib1 then
        begin
        THEdit(GetControl('LIBCRIT2')).Text := 'Agence';
        THEdit(GetControl('LIBCRIT2')).Enabled:=false;
        TToolBarButton97(GetControl('SBCRIT2')).Enabled:=false;
        end
    else
        begin
        THEdit(GetControl('LIBCRIT1')).Text := 'Agence';
        THEdit(GetControl('LIBCRIT1')).Enabled:=false;
        TToolBarButton97(GetControl('SBCRIT1')).Enabled:=false;
        THEdit(GetControl('LIBCRIT2')).Text := RechDom('GCZONELIBRE','CT2',False);
        end
    end
else
//    THEdit(GetControl('LIBCRIT2')).Text := RechDom('GCCATTIERSTABLE','AL2',False);
    THEdit(GetControl('LIBCRIT2')).Text := RechDom('GCZONELIBRE','CT2',False);

(*THEdit(GetControl('LIBCRIT3')).Text := RechDom('GCCATTIERSTABLE','AL3',False);
THEdit(GetControl('LIBCRIT4')).Text := RechDom('GCCATTIERSTABLE','AL4',False);
THEdit(GetControl('LIBCRIT5')).Text := RechDom('GCCATTIERSTABLE','AL5',False);
THEdit(GetControl('LIBCRIT6')).Text := RechDom('GCCATTIERSTABLE','AL6',False);
THEdit(GetControl('LIBCRIT7')).Text := RechDom('GCCATTIERSTABLE','AL7',False);
THEdit(GetControl('LIBCRIT8')).Text := RechDom('GCCATTIERSTABLE','AL8',False);
THEdit(GetControl('LIBCRIT9')).Text := RechDom('GCCATTIERSTABLE','AL9',False);
THEdit(GetControl('LIBCRIT10')).Text := RechDom('GCCATTIERSTABLE','ALA',False);
THEdit(GetControl('LIBMONTANT1')).Text := RechDom('GCCATTIERSTABLE','BL1',False);
THEdit(GetControl('LIBMONTANT2')).Text := RechDom('GCCATTIERSTABLE','BL2',False);
THEdit(GetControl('LIBMONTANT3')).Text := RechDom('GCCATTIERSTABLE','BL3',False);
THEdit(GetControl('LIBDATE1')).Text := RechDom('GCCATTIERSTABLE','CL1',False);
THEdit(GetControl('LIBDATE2')).Text := RechDom('GCCATTIERSTABLE','CL2',False);
THEdit(GetControl('LIBDATE3')).Text := RechDom('GCCATTIERSTABLE','CL3',False);*)
THEdit(GetControl('LIBCRIT3')).Text := RechDom('GCZONELIBRE','CT3',False);
THEdit(GetControl('LIBCRIT4')).Text := RechDom('GCZONELIBRE','CT4',False);
THEdit(GetControl('LIBCRIT5')).Text := RechDom('GCZONELIBRE','CT5',False);
THEdit(GetControl('LIBCRIT6')).Text := RechDom('GCZONELIBRE','CT6',False);
THEdit(GetControl('LIBCRIT7')).Text := RechDom('GCZONELIBRE','CT7',False);
THEdit(GetControl('LIBCRIT8')).Text := RechDom('GCZONELIBRE','CT8',False);
THEdit(GetControl('LIBCRIT9')).Text := RechDom('GCZONELIBRE','CT9',False);
THEdit(GetControl('LIBCRIT10')).Text := RechDom('GCZONELIBRE','CTA',False);
THEdit(GetControl('LIBMONTANT1')).Text := RechDom('GCZONELIBRE','CM1',False);
THEdit(GetControl('LIBMONTANT2')).Text := RechDom('GCZONELIBRE','CM2',False);
THEdit(GetControl('LIBMONTANT3')).Text := RechDom('GCZONELIBRE','CM3',False);
THEdit(GetControl('LIBDATE1')).Text := RechDom('GCZONELIBRE','CD1',False);
THEdit(GetControl('LIBDATE2')).Text := RechDom('GCZONELIBRE','CD2',False);
THEdit(GetControl('LIBDATE3')).Text := RechDom('GCZONELIBRE','CD3',False);


// Libellés ressources
if bRess1 then
    begin
    THEdit(GetControl('LIBRES1')).Text := 'Associé';
    THEdit(GetControl('LIBRES1')).Enabled:=false;
    end
else
//    THEdit(GetControl('LIBRES1')).Text := RechDom('AFCATIERSTABLE','AL1',False);
    THEdit(GetControl('LIBRES1')).Text := RechDom('GCZONELIBRE','CR1',False);


if bRess2 then
    begin
    if bRess1 then
        begin
        THEdit(GetControl('LIBRES2')).Text := 'Chef de groupe';
        THEdit(GetControl('LIBRES2')).Enabled:=false;
        end
    else
        begin
        THEdit(GetControl('LIBRES1')).Text := 'Chef de groupe';
        THEdit(GetControl('LIBRES1')).Enabled:=false;
        THEdit(GetControl('LIBRES2')).Text := RechDom('GCZONELIBRE','CR2',False);
        end;
    end
else
//    THEdit(GetControl('LIBRES2')).Text := RechDom('AFCATIERSTABLE','AL2',False);
    THEdit(GetControl('LIBRES2')).Text := RechDom('GCZONELIBRE','CR2',False);


if bRess3 then
    begin
    if bRess2 then
        begin
        if bRess1 then
            begin
            THEdit(GetControl('LIBRES3')).Text := 'Responsable';
            THEdit(GetControl('LIBRES3')).Enabled:=false;
            end
        else
            begin
            THEdit(GetControl('LIBRES2')).Text := 'Responsable';
            THEdit(GetControl('LIBRES2')).Enabled:=false;
            THEdit(GetControl('LIBRES3')).Text := RechDom('GCZONELIBRE','CR3',False);
            end;
        end
    else
        begin
        if bRess1 then
            begin
            THEdit(GetControl('LIBRES2')).Text := 'Responsable';
            THEdit(GetControl('LIBRES2')).Enabled:=false;
            THEdit(GetControl('LIBRES3')).Text := RechDom('GCZONELIBRE','CR3',False);
            end
        else
            begin
            THEdit(GetControl('LIBRES1')).Text := 'Responsable';
            THEdit(GetControl('LIBRES1')).Enabled:=false;
            THEdit(GetControl('LIBRES3')).Text := RechDom('GCZONELIBRE','CR3',False);
            end;
        end;
    end
else
//    THEdit(GetControl('LIBRES3')).Text := RechDom('AFCATIERSTABLE','AL3',False);
    THEdit(GetControl('LIBRES3')).Text := RechDom('GCZONELIBRE','CR3',False);


// Branchement des boutons sur tablettes criteres
TToolBarButton97(GetControl('SBCRIT1')).OnClick:=bCritere1Click;
TToolBarButton97(GetControl('SBCRIT2')).OnClick:=bCritere2Click;
TToolBarButton97(GetControl('SBCRIT3')).OnClick:=bCritere3Click;
TToolBarButton97(GetControl('SBCRIT4')).OnClick:=bCritere4Click;
TToolBarButton97(GetControl('SBCRIT5')).OnClick:=bCritere5Click;
TToolBarButton97(GetControl('SBCRIT6')).OnClick:=bCritere6Click;
TToolBarButton97(GetControl('SBCRIT7')).OnClick:=bCritere7Click;
TToolBarButton97(GetControl('SBCRIT8')).OnClick:=bCritere8Click;
TToolBarButton97(GetControl('SBCRIT9')).OnClick:=bCritere9Click;
TToolBarButton97(GetControl('SBCRIT10')).OnClick:=bCritere10Click;
end;

procedure TOF_AFASSCREATLIBTIER.OnUpdate;
Begin
inherited;
(*
if Not bCritLib1 then
    ModifierTablette('CC', 'ZLB', 'AL1', THEdit(GetControl('LIBCRIT1')).Text, 3);
if Not bCritLib2 then
    ModifierTablette('CC', 'ZLB', 'AL2', THEdit(GetControl('LIBCRIT2')).Text, 3);
ModifierTablette('CC', 'ZLB', 'AL3', THEdit(GetControl('LIBCRIT3')).Text, 3);
ModifierTablette('CC', 'ZLB', 'AL4', THEdit(GetControl('LIBCRIT4')).Text, 3);
ModifierTablette('CC', 'ZLB', 'AL5', THEdit(GetControl('LIBCRIT5')).Text, 3);
ModifierTablette('CC', 'ZLB', 'AL6', THEdit(GetControl('LIBCRIT6')).Text, 3);
ModifierTablette('CC', 'ZLB', 'AL7', THEdit(GetControl('LIBCRIT7')).Text, 3);
ModifierTablette('CC', 'ZLB', 'AL8', THEdit(GetControl('LIBCRIT8')).Text, 3);
ModifierTablette('CC', 'ZLB', 'AL9', THEdit(GetControl('LIBCRIT9')).Text, 3);
ModifierTablette('CC', 'ZLB', 'ALA', THEdit(GetControl('LIBCRIT10')).Text, 3);
ModifierTablette('CC', 'ZLB', 'BL1', THEdit(GetControl('LIBMONTANT1')).Text, 3);
ModifierTablette('CC', 'ZLB', 'BL2', THEdit(GetControl('LIBMONTANT2')).Text, 3);
ModifierTablette('CC', 'ZLB', 'BL3', THEdit(GetControl('LIBMONTANT3')).Text, 3);
ModifierTablette('CC', 'ZLB', 'CL1', THEdit(GetControl('LIBDATE1')).Text, 3);
ModifierTablette('CC', 'ZLB', 'CL2', THEdit(GetControl('LIBDATE2')).Text, 3);
ModifierTablette('CC', 'ZLB', 'CL3', THEdit(GetControl('LIBDATE3')).Text, 3);
AvertirTable('GCCATTIERSTABLE');

// Ressources
if Not bRess1 then
    ModifierTablette('CC', 'ZLF', 'AL1', THEdit(GetControl('LIBRES1')).Text, 3);
if Not bRess2 then
    ModifierTablette('CC', 'ZLF', 'AL2', THEdit(GetControl('LIBRES2')).Text, 3);
if Not bRess3 then
    ModifierTablette('CC', 'ZLF', 'AL3', THEdit(GetControl('LIBRES3')).Text, 3);
AvertirTable('AFCATIERSTABLE');
*)

if Not bCritLib1 then
    ModifierTablette('CC', 'ZLI', 'CT1', THEdit(GetControl('LIBCRIT1')).Text, 3);
if Not bCritLib2 then
    ModifierTablette('CC', 'ZLI', 'CT2', THEdit(GetControl('LIBCRIT2')).Text, 3);
ModifierTablette('CC', 'ZLI', 'CT3', THEdit(GetControl('LIBCRIT3')).Text, 3);
ModifierTablette('CC', 'ZLI', 'CT4', THEdit(GetControl('LIBCRIT4')).Text, 3);
ModifierTablette('CC', 'ZLI', 'CT5', THEdit(GetControl('LIBCRIT5')).Text, 3);
ModifierTablette('CC', 'ZLI', 'CT6', THEdit(GetControl('LIBCRIT6')).Text, 3);
ModifierTablette('CC', 'ZLI', 'CT7', THEdit(GetControl('LIBCRIT7')).Text, 3);
ModifierTablette('CC', 'ZLI', 'CT8', THEdit(GetControl('LIBCRIT8')).Text, 3);
ModifierTablette('CC', 'ZLI', 'CT9', THEdit(GetControl('LIBCRIT9')).Text, 3);
ModifierTablette('CC', 'ZLI', 'CTA', THEdit(GetControl('LIBCRIT10')).Text, 3);
ModifierTablette('CC', 'ZLI', 'CM1', THEdit(GetControl('LIBMONTANT1')).Text, 3);
ModifierTablette('CC', 'ZLI', 'CM2', THEdit(GetControl('LIBMONTANT2')).Text, 3);
ModifierTablette('CC', 'ZLI', 'CM3', THEdit(GetControl('LIBMONTANT3')).Text, 3);
ModifierTablette('CC', 'ZLI', 'CD1', THEdit(GetControl('LIBDATE1')).Text, 3);
ModifierTablette('CC', 'ZLI', 'CD2', THEdit(GetControl('LIBDATE2')).Text, 3);
ModifierTablette('CC', 'ZLI', 'CD3', THEdit(GetControl('LIBDATE3')).Text, 3);
AvertirTable('GCZONELIBRE');

// Ressources
if Not bRess1 then
    ModifierTablette('CC', 'ZLI', 'CR1', THEdit(GetControl('LIBRES1')).Text, 3);
if Not bRess2 then
    ModifierTablette('CC', 'ZLI', 'CR2', THEdit(GetControl('LIBRES2')).Text, 3);
if Not bRess3 then
    ModifierTablette('CC', 'ZLI', 'CR3', THEdit(GetControl('LIBRES3')).Text, 3);
AvertirTable('GCZONELIBRE');

End;

procedure TOF_AFASSCREATLIBTIER.bCritere1Click(Sender: TObject);
begin
ParamTable ('GCLIBRETIERS1',taCreat,0,nil, 6, THEdit(GetControl('LIBCRIT1')).Text)
end;

procedure TOF_AFASSCREATLIBTIER.bCritere2Click(Sender: TObject);
begin
ParamTable ('GCLIBRETIERS2',taCreat,0,nil, 6, THEdit(GetControl('LIBCRIT2')).Text)
end;

procedure TOF_AFASSCREATLIBTIER.bCritere3Click(Sender: TObject);
begin
ParamTable ('GCLIBRETIERS3',taCreat,0,nil, 6, THEdit(GetControl('LIBCRIT3')).Text)
end;

procedure TOF_AFASSCREATLIBTIER.bCritere4Click(Sender: TObject);
begin
ParamTable ('GCLIBRETIERS4',taCreat,0,nil, 6, THEdit(GetControl('LIBCRIT4')).Text)
end;

procedure TOF_AFASSCREATLIBTIER.bCritere5Click(Sender: TObject);
begin
ParamTable ('GCLIBRETIERS5',taCreat,0,nil, 6, THEdit(GetControl('LIBCRIT5')).Text)
end;

procedure TOF_AFASSCREATLIBTIER.bCritere6Click(Sender: TObject);
begin
ParamTable ('GCLIBRETIERS6',taCreat,0,nil, 6, THEdit(GetControl('LIBCRIT6')).Text)
end;

procedure TOF_AFASSCREATLIBTIER.bCritere7Click(Sender: TObject);
begin
ParamTable ('GCLIBRETIERS7',taCreat,0,nil, 6, THEdit(GetControl('LIBCRIT7')).Text)
end;

procedure TOF_AFASSCREATLIBTIER.bCritere8Click(Sender: TObject);
begin
ParamTable ('GCLIBRETIERS8',taCreat,0,nil, 6, THEdit(GetControl('LIBCRIT8')).Text)
end;

procedure TOF_AFASSCREATLIBTIER.bCritere9Click(Sender: TObject);
begin
ParamTable ('GCLIBRETIERS9',taCreat,0,nil, 6, THEdit(GetControl('LIBCRIT9')).Text)
end;

procedure TOF_AFASSCREATLIBTIER.bCritere10Click(Sender: TObject);
begin
ParamTable ('GCLIBRETIERS10',taCreat,0,nil, 6, THEdit(GetControl('LIBCRIT10')).Text)
end;

Function AFLanceFiche_CreatLibelClt(Argument:string):variant;
begin
result:=AGLLanceFiche ('AFF','AFASSCREATLIBTIER','','',Argument);
end;

Initialization
registerclasses([TOF_AFASSCREATLIBTIER]);
end.
