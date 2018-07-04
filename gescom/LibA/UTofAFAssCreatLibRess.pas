unit UTofAFAssCreatLibRess;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls, 
      HCtrls,Ent1,HMsgBox,UTOF,DicoAF,SaisUtil,EntGC,HTB97,
{$IFDEF EAGLCLIENT}
       Maineagl, Etablette, Emul,
{$ELSE}
        db, dbTables,HDB,DBCtrls,FE_Main, tablette, mul  ,
{$ENDIF}
      Buttons, ExtCtrls, HEnt1, utob, M3FP;

Type
     TOF_AFASSCREATLIBRESS = Class (TOF)
         procedure OnArgument(stArgument : String ) ; override ;
         procedure OnUpdate ; override ;
         procedure OnClose ; override ;
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
        source:string;
        CritLib:boolean;
        bCritLib1:boolean;
        bCritLib2:boolean;
        bCritLib3:boolean;
        sValCrit1:string;
        sValCrit2:string;
        sValCrit3:string;
     END ;

Function AFLanceFiche_CreatLibelRess(Argument:string):variant;

implementation


procedure TOF_AFASSCREATLIBRESS.OnArgument(stArgument : String );
var
Crit:string;
haut:integer;
champ, valeur{, value} : string;
x:integer;
Begin
Inherited;
Source:='';
CritLib:=false;
bCritLib1:=false;
bCritLib2:=false;
bCritLib3:=false;

Crit:=(Trim(ReadTokenSt(stArgument)));
while (Crit<>'') do
    begin
    Champ:=Crit;
    X:=pos('=',Crit);
    if x<>0 then
           begin
           Champ:=copy(Crit,1,X-1);
           Valeur:=Copy (Crit,X+1,length(Crit)-X);
           end;

    if (Champ='CRIT1') then
        begin
        bCritLib1:=true;
        sValCrit1 := Valeur;
        end
    else
    if (Champ='CRIT2') then
        begin
        bCritLib2:=true;
        sValCrit2 := Valeur;
        end
    else
    if (Champ='CRIT3') then
        begin
        bCritLib3:=true;
        sValCrit3 := Valeur;
        end
    else
    if (Champ='RES') or (Champ='AFF') then Source:=Crit
    else
    if (Champ='CRITERES') then CritLib:=True;

    Crit:=(Trim(ReadTokenSt(stArgument)));
    end;

if CritLib then TGroupBox(GetControl('GBLIBELLESLIBRES')).visible:=true
            else
                begin
                haut:= TGroupBox(GetControl('GBTEXTESLIBRES')).Height;
                TGroupBox(GetControl('GBLIBELLESLIBRES')).visible:=false;
                Ecran.Height := Ecran.Height-TGroupBox(GetControl('GBLIBELLESLIBRES')).Height;
                TGroupBox(GetControl('GBTEXTESLIBRES')).top:= TGroupBox(GetControl('GBLIBELLESLIBRES')).top;
                TGroupBox(GetControl('GBDATES')).top:= TGroupBox(GetControl('GBLIBELLESLIBRES')).top;
                TGroupBox(GetControl('GBBOOLEENS')).top:= TGroupBox(GetControl('GBTEXTESLIBRES')).top + haut +5;
                TGroupBox(GetControl('GBMONTANTS')).top:= TGroupBox(GetControl('GBTEXTESLIBRES')).top + haut +5;
                TGroupBox(GetControl('GBTEXTESLIBRES')).Height:=haut;
                TGroupBox(GetControl('GBDATES')).Height:=haut;
                TGroupBox(GetControl('GBBOOLEENS')).Height:=haut;
                TGroupBox(GetControl('GBMONTANTS')).Height:=haut;
                end;

if (source = 'RES') then
    begin
    if bCritLib1 then
        begin
        THEdit(GetControl('CRITERELIB1')).Text := sValCrit1;
        THEdit(GetControl('CRITERELIB1')).Enabled:=false;
        TToolBarButton97(GetControl('BCRITERE1')).Enabled:=false;
        end
    else
        THEdit(GetControl('CRITERELIB1')).Text := RechDom('GCZONELIBRE','RT1',False);
(*
    if bCritLib2 then
        begin
        if bCritLib1 then
            begin
            THEdit(GetControl('CRITERELIB2')).Text := sValCrit2;
            THEdit(GetControl('CRITERELIB2')).Enabled:=false;
            TToolBarButton97(GetControl('BCRITERE2')).Enabled:=false;
            end
        else
            begin
            THEdit(GetControl('CRITERELIB1')).Text := sValCrit2;
            THEdit(GetControl('CRITERELIB1')).Enabled:=false;
            TToolBarButton97(GetControl('BCRITERE1')).Enabled:=false;
            THEdit(GetControl('CRITERELIB2')).Text := RechDom('GCZONELIBRE','RT2',False);
            end
        end
    else
        THEdit(GetControl('CRITERELIB2')).Text := RechDom('GCZONELIBRE','RT2',False);


    if bCritLib3 then
        begin
        if bCritLib2 then
            begin
            if bCritLib1 then
                begin
                THEdit(GetControl('CRITERELIB3')).Text := sValCrit3;
                THEdit(GetControl('CRITERELIB3')).Enabled:=false;
                TToolBarButton97(GetControl('BCRITERE3')).Enabled:=false;
                end
            else
                begin
                THEdit(GetControl('CRITERELIB2')).Text := sValCrit3;
                THEdit(GetControl('CRITERELIB2')).Enabled:=false;
                TToolBarButton97(GetControl('BCRITERE2')).Enabled:=false;
                THEdit(GetControl('CRITERELIB3')).Text := RechDom('GCZONELIBRE','RT3',False);
                end;
            end
        else
            begin
            if bCritLib1 then
                begin
                THEdit(GetControl('CRITERELIB2')).Text := sValCrit3;
                THEdit(GetControl('CRITERELIB2')).Enabled:=false;
                TToolBarButton97(GetControl('BCRITERE2')).Enabled:=false;
                THEdit(GetControl('CRITERELIB3')).Text := RechDom('GCZONELIBRE','RT3',False);
                end
            else
                begin
                THEdit(GetControl('CRITERELIB1')).Text := sValCrit3;
                THEdit(GetControl('CRITERELIB1')).Enabled:=false;
                TToolBarButton97(GetControl('BCRITERE1')).Enabled:=false;
                THEdit(GetControl('CRITERELIB3')).Text := RechDom('GCZONELIBRE','RT3',False);
                end;
            end;
        end
    else
        THEdit(GetControl('CRITERELIB3')).Text := RechDom('GCZONELIBRE','RT3',False);
 *)
(*    if CritLib then
        begin
        THEdit(GetControl('CRITERELIB1')).Text := RechDom('AFCARESTABLE','AL1',False);
        THEdit(GetControl('CRITERELIB2')).Text := RechDom('AFCARESTABLE','AL2',False);
        THEdit(GetControl('CRITERELIB3')).Text := RechDom('AFCARESTABLE','AL3',False);
        end;
    THEdit(GetControl('LIBBOOLEEN1')).Text := RechDom('AFCARESTABLE','BL1',False);
    THEdit(GetControl('LIBBOOLEEN2')).Text := RechDom('AFCARESTABLE','BL2',False);
    THEdit(GetControl('LIBBOOLEEN3')).Text := RechDom('AFCARESTABLE','BL3',False);
    THEdit(GetControl('LIBTEXTLIB1')).Text := RechDom('AFCARESTABLE','CL1',False);
    THEdit(GetControl('LIBTEXTLIB2')).Text := RechDom('AFCARESTABLE','CL2',False);
    THEdit(GetControl('LIBTEXTLIB3')).Text := RechDom('AFCARESTABLE','CL3',False);
    THEdit(GetControl('LIBDATE1')).Text := RechDom('AFCARESTABLE','DL1',False);
    THEdit(GetControl('LIBDATE2')).Text := RechDom('AFCARESTABLE','DL2',False);
    THEdit(GetControl('LIBDATE3')).Text := RechDom('AFCARESTABLE','DL3',False);
    THEdit(GetControl('LIBMONTANT1')).Text := RechDom('AFCARESTABLE','EL1',False);
    THEdit(GetControl('LIBMONTANT2')).Text := RechDom('AFCARESTABLE','EL2',False);
    THEdit(GetControl('LIBMONTANT3')).Text := RechDom('AFCARESTABLE','EL3',False);  *)
    if CritLib then
        begin
        THEdit(GetControl('CRITERELIB1')).Text := RechDom('GCZONELIBRE','RT1',False);
        THEdit(GetControl('CRITERELIB2')).Text := RechDom('GCZONELIBRE','RT2',False);
        THEdit(GetControl('CRITERELIB3')).Text := RechDom('GCZONELIBRE','RT3',False);
        THEdit(GetControl('CRITERELIB4')).Text := RechDom('GCZONELIBRE','RT4',False);
        THEdit(GetControl('CRITERELIB5')).Text := RechDom('GCZONELIBRE','RT5',False);
        THEdit(GetControl('CRITERELIB6')).Text := RechDom('GCZONELIBRE','RT6',False);
        THEdit(GetControl('CRITERELIB7')).Text := RechDom('GCZONELIBRE','RT7',False);
        THEdit(GetControl('CRITERELIB8')).Text := RechDom('GCZONELIBRE','RT8',False);
        THEdit(GetControl('CRITERELIB9')).Text := RechDom('GCZONELIBRE','RT9',False);
        THEdit(GetControl('CRITERELIB10')).Text := RechDom('GCZONELIBRE','RTA',False);
        end;
    THEdit(GetControl('LIBBOOLEEN1')).Text := RechDom('GCZONELIBRE','RB1',False);
    THEdit(GetControl('LIBBOOLEEN2')).Text := RechDom('GCZONELIBRE','RB2',False);
    THEdit(GetControl('LIBBOOLEEN3')).Text := RechDom('GCZONELIBRE','RB3',False);
    THEdit(GetControl('LIBTEXTLIB1')).Text := RechDom('GCZONELIBRE','RC1',False);
    THEdit(GetControl('LIBTEXTLIB2')).Text := RechDom('GCZONELIBRE','RC2',False);
    THEdit(GetControl('LIBTEXTLIB3')).Text := RechDom('GCZONELIBRE','RC3',False);
    THEdit(GetControl('LIBDATE1')).Text := RechDom('GCZONELIBRE','RD1',False);
    THEdit(GetControl('LIBDATE2')).Text := RechDom('GCZONELIBRE','RD2',False);
    THEdit(GetControl('LIBDATE3')).Text := RechDom('GCZONELIBRE','RD3',False);
    THEdit(GetControl('LIBMONTANT1')).Text := RechDom('GCZONELIBRE','RM1',False);
    THEdit(GetControl('LIBMONTANT2')).Text := RechDom('GCZONELIBRE','RM2',False);
    THEdit(GetControl('LIBMONTANT3')).Text := RechDom('GCZONELIBRE','RM3',False);

    Ecran.caption:= Ecran.caption + TraduitGA(' de la ressource');
    end
else
if (source = 'AFF') then
    begin
(*    if CritLib then
        begin
        THEdit(GetControl('CRITERELIB1')).Text := RechDom('AFCAAFFTABLE','AL1',False);
        THEdit(GetControl('CRITERELIB2')).Text := RechDom('AFCAAFFTABLE','AL2',False);
        THEdit(GetControl('CRITERELIB3')).Text := RechDom('AFCAAFFTABLE','AL3',False);
        end;
    THEdit(GetControl('LIBBOOLEEN1')).Text := RechDom('AFCAAFFTABLE','BL1',False);
    THEdit(GetControl('LIBBOOLEEN2')).Text := RechDom('AFCAAFFTABLE','BL2',False);
    THEdit(GetControl('LIBBOOLEEN3')).Text := RechDom('AFCAAFFTABLE','BL3',False);
    THEdit(GetControl('LIBTEXTLIB1')).Text := RechDom('AFCAAFFTABLE','CL1',False);
    THEdit(GetControl('LIBTEXTLIB2')).Text := RechDom('AFCAAFFTABLE','CL2',False);
    THEdit(GetControl('LIBTEXTLIB3')).Text := RechDom('AFCAAFFTABLE','CL3',False);
    THEdit(GetControl('LIBDATE1')).Text := RechDom('AFCAAFFTABLE','DL1',False);
    THEdit(GetControl('LIBDATE2')).Text := RechDom('AFCAAFFTABLE','DL2',False);
    THEdit(GetControl('LIBDATE3')).Text := RechDom('AFCAAFFTABLE','DL3',False);
    THEdit(GetControl('LIBMONTANT1')).Text := RechDom('AFCAAFFTABLE','EL1',False);
    THEdit(GetControl('LIBMONTANT2')).Text := RechDom('AFCAAFFTABLE','EL2',False);
    THEdit(GetControl('LIBMONTANT3')).Text := RechDom('AFCAAFFTABLE','EL3',False);*)
    if CritLib then
        begin
        THEdit(GetControl('CRITERELIB1')).Text := RechDom('GCZONELIBRE','MT1',False);
        THEdit(GetControl('CRITERELIB2')).Text := RechDom('GCZONELIBRE','MT2',False);
        THEdit(GetControl('CRITERELIB3')).Text := RechDom('GCZONELIBRE','MT3',False);
        THEdit(GetControl('CRITERELIB4')).Text := RechDom('GCZONELIBRE','MT4',False);
        THEdit(GetControl('CRITERELIB5')).Text := RechDom('GCZONELIBRE','MT5',False);
        THEdit(GetControl('CRITERELIB6')).Text := RechDom('GCZONELIBRE','MT6',False);
        THEdit(GetControl('CRITERELIB7')).Text := RechDom('GCZONELIBRE','MT7',False);
        THEdit(GetControl('CRITERELIB8')).Text := RechDom('GCZONELIBRE','MT8',False);
        THEdit(GetControl('CRITERELIB9')).Text := RechDom('GCZONELIBRE','MT9',False);
        THEdit(GetControl('CRITERELIB10')).Text := RechDom('GCZONELIBRE','MTA',False);
        end;
    THEdit(GetControl('LIBBOOLEEN1')).Text := RechDom('GCZONELIBRE','MB1',False);
    THEdit(GetControl('LIBBOOLEEN2')).Text := RechDom('GCZONELIBRE','MB2',False);
    THEdit(GetControl('LIBBOOLEEN3')).Text := RechDom('GCZONELIBRE','MB3',False);
    THEdit(GetControl('LIBTEXTLIB1')).Text := RechDom('GCZONELIBRE','MC1',False);
    THEdit(GetControl('LIBTEXTLIB2')).Text := RechDom('GCZONELIBRE','MC2',False);
    THEdit(GetControl('LIBTEXTLIB3')).Text := RechDom('GCZONELIBRE','MC3',False);
    THEdit(GetControl('LIBDATE1')).Text := RechDom('GCZONELIBRE','MD1',False);
    THEdit(GetControl('LIBDATE2')).Text := RechDom('GCZONELIBRE','MD2',False);
    THEdit(GetControl('LIBDATE3')).Text := RechDom('GCZONELIBRE','MD3',False);
    THEdit(GetControl('LIBMONTANT1')).Text := RechDom('GCZONELIBRE','MM1',False);
    THEdit(GetControl('LIBMONTANT2')).Text := RechDom('GCZONELIBRE','MM2',False);
    THEdit(GetControl('LIBMONTANT3')).Text := RechDom('GCZONELIBRE','MM3',False);

    Ecran.caption:= Ecran.caption + TraduitGA(' de l''affaire');
    end;

if CritLib then
    begin
    TToolBarButton97(GetControl('BCRITERE1')).OnClick:=bCritere1Click;
    TToolBarButton97(GetControl('BCRITERE2')).OnClick:=bCritere2Click;
    TToolBarButton97(GetControl('BCRITERE3')).OnClick:=bCritere3Click;
    TToolBarButton97(GetControl('BCRITERE4')).OnClick:=bCritere4Click;
    TToolBarButton97(GetControl('BCRITERE5')).OnClick:=bCritere5Click;
    TToolBarButton97(GetControl('BCRITERE6')).OnClick:=bCritere6Click;
    TToolBarButton97(GetControl('BCRITERE7')).OnClick:=bCritere7Click;
    TToolBarButton97(GetControl('BCRITERE8')).OnClick:=bCritere8Click;
    TToolBarButton97(GetControl('BCRITERE9')).OnClick:=bCritere9Click;
    TToolBarButton97(GetControl('BCRITERE10')).OnClick:=bCritere10Click;
    end;

updatecaption(Ecran);
end;

procedure TOF_AFASSCREATLIBRESS.OnClose;
//var crit:string;
Begin
 Inherited;
(*Retour:='';
crit := THEdit(GetControl('CRITERELIB1')).Text;
if (crit <>'') and (Uppercase(crit) <> 'TABLE LIBRE 1') then
    Retour:='CRIT1=' + crit + ';';

crit := THEdit(GetControl('CRITERELIB2')).Text;
if (crit <>'') and (Uppercase(crit) <> 'TABLE LIBRE 2') then
    Retour:='CRIT2=' + crit + ';';

crit := THEdit(GetControl('CRITERELIB3')).Text;
if (crit <>'') and (Uppercase(crit) <> 'TABLE LIBRE 3') then
    Retour:='CRIT3=' + crit + ';';
*)
end;

procedure TOF_AFASSCREATLIBRESS.bCritere1Click(Sender: TObject);
begin
if (source = 'RES') then
    ParamTable ('AFTLIBRERES1',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB1')).Text)
else
if (source = 'AFF') then
    ParamTable ('AFTLIBREAFF1',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB1')).Text)
;
end;

procedure TOF_AFASSCREATLIBRESS.bCritere2Click(Sender: TObject);
begin
if (source = 'RES') then
    ParamTable ('AFTLIBRERES2',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB2')).Text)
else
if (source = 'AFF') then
    ParamTable ('AFTLIBREAFF2',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB2')).Text)
;
end;

procedure TOF_AFASSCREATLIBRESS.bCritere3Click(Sender: TObject);
begin
if (source = 'RES') then
    ParamTable ('AFTLIBRERES3',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB3')).Text)
else
if (source = 'AFF') then
    ParamTable ('AFTLIBREAFF3',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB3')).Text)
;
end;

procedure TOF_AFASSCREATLIBRESS.bCritere4Click(Sender: TObject);
begin
if (source = 'RES') then
    ParamTable ('AFTLIBRERES4',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB4')).Text)
else
if (source = 'AFF') then
    ParamTable ('AFTLIBREAFF4',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB4')).Text)
;
end;

procedure TOF_AFASSCREATLIBRESS.bCritere5Click(Sender: TObject);
begin
if (source = 'RES') then
    ParamTable ('AFTLIBRERES5',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB5')).Text)
else
if (source = 'AFF') then
    ParamTable ('AFTLIBREAFF5',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB5')).Text)
;
end;

procedure TOF_AFASSCREATLIBRESS.bCritere6Click(Sender: TObject);
begin
if (source = 'RES') then
    ParamTable ('AFTLIBRERES6',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB6')).Text)
else
if (source = 'AFF') then
    ParamTable ('AFTLIBREAFF6',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB6')).Text)
;
end;

procedure TOF_AFASSCREATLIBRESS.bCritere7Click(Sender: TObject);
begin
if (source = 'RES') then
    ParamTable ('AFTLIBRERES7',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB7')).Text)
else
if (source = 'AFF') then
    ParamTable ('AFTLIBREAFF7',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB7')).Text)
;
end;

procedure TOF_AFASSCREATLIBRESS.bCritere8Click(Sender: TObject);
begin
if (source = 'RES') then
    ParamTable ('AFTLIBRERES8',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB8')).Text)
else
if (source = 'AFF') then
    ParamTable ('AFTLIBREAFF8',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB8')).Text)
;
end;

procedure TOF_AFASSCREATLIBRESS.bCritere9Click(Sender: TObject);
begin
if (source = 'RES') then
    ParamTable ('AFTLIBRERES9',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB9')).Text)
else
if (source = 'AFF') then
    ParamTable ('AFTLIBREAFF9',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB9')).Text)
;
end;

procedure TOF_AFASSCREATLIBRESS.bCritere10Click(Sender: TObject);
begin
if (source = 'RES') then
    ParamTable ('AFTLIBRERESA',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB10')).Text)
else
if (source = 'AFF') then
    ParamTable ('AFTLIBREAFFA',taCreat,0,nil, 6, THEdit(GetControl('CRITERELIB10')).Text)
;
end;

procedure TOF_AFASSCREATLIBRESS.OnUpdate;
Begin
inherited;

if (source = 'RES') then
    begin
(*    if CritLib then
        begin
        ModifierTablette('CC', 'ZLD', 'AL1', THEdit(GetControl('CRITERELIB1')).Text, 3);
        ModifierTablette('CC', 'ZLD', 'AL2', THEdit(GetControl('CRITERELIB2')).Text, 3);
        ModifierTablette('CC', 'ZLD', 'AL3', THEdit(GetControl('CRITERELIB3')).Text, 3);
        end;
    ModifierTablette('CC', 'ZLD', 'BL1', THEdit(GetControl('LIBBOOLEEN1')).Text, 3);
    ModifierTablette('CC', 'ZLD', 'BL2', THEdit(GetControl('LIBBOOLEEN2')).Text, 3);
    ModifierTablette('CC', 'ZLD', 'BL3', THEdit(GetControl('LIBBOOLEEN3')).Text, 3);
    ModifierTablette('CC', 'ZLD', 'CL1', THEdit(GetControl('LIBTEXTLIB1')).Text, 3);
    ModifierTablette('CC', 'ZLD', 'CL2', THEdit(GetControl('LIBTEXTLIB2')).Text, 3);
    ModifierTablette('CC', 'ZLD', 'CL3', THEdit(GetControl('LIBTEXTLIB3')).Text, 3);
    ModifierTablette('CC', 'ZLD', 'DL1', THEdit(GetControl('LIBDATE1')).Text, 3);
    ModifierTablette('CC', 'ZLD', 'DL2', THEdit(GetControl('LIBDATE2')).Text, 3);
    ModifierTablette('CC', 'ZLD', 'DL3', THEdit(GetControl('LIBDATE3')).Text, 3);
    ModifierTablette('CC', 'ZLD', 'EL1', THEdit(GetControl('LIBMONTANT1')).Text, 3);
    ModifierTablette('CC', 'ZLD', 'EL2', THEdit(GetControl('LIBMONTANT2')).Text, 3);
    ModifierTablette('CC', 'ZLD', 'EL3', THEdit(GetControl('LIBMONTANT3')).Text, 3);
    AvertirTable('AFCARESTABLE');*)
    if CritLib then
        begin
        (*if Not bCritLib1 then
            ModifierTablette('CC', 'ZLI', 'RT1', THEdit(GetControl('CRITERELIB1')).Text, 3);
        if Not bCritLib2 then
            ModifierTablette('CC', 'ZLI', 'RT2', THEdit(GetControl('CRITERELIB2')).Text, 3);
        if Not bCritLib3 then
            ModifierTablette('CC', 'ZLI', 'RT3', THEdit(GetControl('CRITERELIB3')).Text, 3);
            *)
        ModifierTablette('CC', 'ZLI', 'RT1', THEdit(GetControl('CRITERELIB1')).Text, 3);
        ModifierTablette('CC', 'ZLI', 'RT2', THEdit(GetControl('CRITERELIB2')).Text, 3);
        ModifierTablette('CC', 'ZLI', 'RT3', THEdit(GetControl('CRITERELIB3')).Text, 3);
        ModifierTablette('CC', 'ZLI', 'RT4', THEdit(GetControl('CRITERELIB4')).Text, 3);
        ModifierTablette('CC', 'ZLI', 'RT5', THEdit(GetControl('CRITERELIB5')).Text, 3);
        ModifierTablette('CC', 'ZLI', 'RT6', THEdit(GetControl('CRITERELIB6')).Text, 3);
        ModifierTablette('CC', 'ZLI', 'RT7', THEdit(GetControl('CRITERELIB7')).Text, 3);
        ModifierTablette('CC', 'ZLI', 'RT8', THEdit(GetControl('CRITERELIB8')).Text, 3);
        ModifierTablette('CC', 'ZLI', 'RT9', THEdit(GetControl('CRITERELIB9')).Text, 3);
        ModifierTablette('CC', 'ZLI', 'RTA', THEdit(GetControl('CRITERELIB10')).Text, 3);
        end;
    ModifierTablette('CC', 'ZLI', 'RB1', THEdit(GetControl('LIBBOOLEEN1')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'RB2', THEdit(GetControl('LIBBOOLEEN2')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'RB3', THEdit(GetControl('LIBBOOLEEN3')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'RC1', THEdit(GetControl('LIBTEXTLIB1')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'RC2', THEdit(GetControl('LIBTEXTLIB2')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'RC3', THEdit(GetControl('LIBTEXTLIB3')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'RD1', THEdit(GetControl('LIBDATE1')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'RD2', THEdit(GetControl('LIBDATE2')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'RD3', THEdit(GetControl('LIBDATE3')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'RM1', THEdit(GetControl('LIBMONTANT1')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'RM2', THEdit(GetControl('LIBMONTANT2')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'RM3', THEdit(GetControl('LIBMONTANT3')).Text, 3);
    AvertirTable('GCZONELIBRE');
    end
else
if (source = 'AFF') then
    begin
(*    if CritLib then
        begin
        ModifierTablette('CC', 'ZLE', 'AL1', THEdit(GetControl('CRITERELIB1')).Text, 3);
        ModifierTablette('CC', 'ZLE', 'AL2', THEdit(GetControl('CRITERELIB2')).Text, 3);
        ModifierTablette('CC', 'ZLE', 'AL3', THEdit(GetControl('CRITERELIB3')).Text, 3);
        end;
    ModifierTablette('CC', 'ZLE', 'BL1', THEdit(GetControl('LIBBOOLEEN1')).Text, 3);
    ModifierTablette('CC', 'ZLE', 'BL2', THEdit(GetControl('LIBBOOLEEN2')).Text, 3);
    ModifierTablette('CC', 'ZLE', 'BL3', THEdit(GetControl('LIBBOOLEEN3')).Text, 3);
    ModifierTablette('CC', 'ZLE', 'CL1', THEdit(GetControl('LIBTEXTLIB1')).Text, 3);
    ModifierTablette('CC', 'ZLE', 'CL2', THEdit(GetControl('LIBTEXTLIB2')).Text, 3);
    ModifierTablette('CC', 'ZLE', 'CL3', THEdit(GetControl('LIBTEXTLIB3')).Text, 3);
    ModifierTablette('CC', 'ZLE', 'DL1', THEdit(GetControl('LIBDATE1')).Text, 3);
    ModifierTablette('CC', 'ZLE', 'DL2', THEdit(GetControl('LIBDATE2')).Text, 3);
    ModifierTablette('CC', 'ZLE', 'DL3', THEdit(GetControl('LIBDATE3')).Text, 3);
    ModifierTablette('CC', 'ZLE', 'EL1', THEdit(GetControl('LIBMONTANT1')).Text, 3);
    ModifierTablette('CC', 'ZLE', 'EL2', THEdit(GetControl('LIBMONTANT2')).Text, 3);
    ModifierTablette('CC', 'ZLE', 'EL3', THEdit(GetControl('LIBMONTANT3')).Text, 3);
    AvertirTable('AFCAAFFTABLE');*)
    if CritLib then
        begin
        ModifierTablette('CC', 'ZLI', 'MT1', THEdit(GetControl('CRITERELIB1')).Text, 3);
        ModifierTablette('CC', 'ZLI', 'MT2', THEdit(GetControl('CRITERELIB2')).Text, 3);
        ModifierTablette('CC', 'ZLI', 'MT3', THEdit(GetControl('CRITERELIB3')).Text, 3);
        ModifierTablette('CC', 'ZLI', 'MT4', THEdit(GetControl('CRITERELIB4')).Text, 3);
        ModifierTablette('CC', 'ZLI', 'MT5', THEdit(GetControl('CRITERELIB5')).Text, 3);
        ModifierTablette('CC', 'ZLI', 'MT6', THEdit(GetControl('CRITERELIB6')).Text, 3);
        ModifierTablette('CC', 'ZLI', 'MT7', THEdit(GetControl('CRITERELIB7')).Text, 3);
        ModifierTablette('CC', 'ZLI', 'MT8', THEdit(GetControl('CRITERELIB8')).Text, 3);
        ModifierTablette('CC', 'ZLI', 'MT9', THEdit(GetControl('CRITERELIB9')).Text, 3);
        ModifierTablette('CC', 'ZLI', 'MTA', THEdit(GetControl('CRITERELIB10')).Text, 3);
        end;
    ModifierTablette('CC', 'ZLI', 'MB1', THEdit(GetControl('LIBBOOLEEN1')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'MB2', THEdit(GetControl('LIBBOOLEEN2')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'MB3', THEdit(GetControl('LIBBOOLEEN3')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'MC1', THEdit(GetControl('LIBTEXTLIB1')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'MC2', THEdit(GetControl('LIBTEXTLIB2')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'MC3', THEdit(GetControl('LIBTEXTLIB3')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'MD1', THEdit(GetControl('LIBDATE1')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'MD2', THEdit(GetControl('LIBDATE2')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'MD3', THEdit(GetControl('LIBDATE3')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'MM1', THEdit(GetControl('LIBMONTANT1')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'MM2', THEdit(GetControl('LIBMONTANT2')).Text, 3);
    ModifierTablette('CC', 'ZLI', 'MM3', THEdit(GetControl('LIBMONTANT3')).Text, 3);
    AvertirTable('GCZONELIBRE');
    end;


End;

Function AFLanceFiche_CreatLibelRess(Argument:string):variant;
begin
result:=AGLLanceFiche ('AFF','AFASSCREATLIBRESS','','',Argument);
end;


Initialization
registerclasses([TOF_AFASSCREATLIBRESS]);
end.
