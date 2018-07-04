unit UTofDispoEtat;

interface

uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls, graphics,
      HCtrls,HEnt1,HMsgBox,UTOF, UTOM,AglInit,Dialogs,
      M3FP, EntGC,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      Fiche, mul, db,DBGrids,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$IFDEF V530}
      EdtEtat,
{$ELSE}
      EdtREtat,
{$ENDIF}      
{$ENDIF}


      grids,UtilArticle,UtilGc ;


Type
     Tof_DispoEtat = Class (TOF)
        procedure OnArgument(Arguments : String ) ; override ;
     END ;

implementation


procedure Tof_DispoEtat.OnArgument(Arguments : String) ;
var iCol : integer ;
begin
Inherited ;
// Paramètrage de la collection
ChangeLibre2('TGA_COLLECTION',Ecran);
// Paramétrage des libellés des familles, stat. article et dimensions
for iCol:=1 to 3 do ChangeLibre2('TGA_FAMILLENIV'+IntToStr(iCol),Ecran);
if (ctxMode in V_PGI.PGIContexte) and (GetPresentation=ART_ORLI) then
  for iCol:=4 to 8 do ChangeLibre2('TGA2_FAMILLENIV'+IntToStr(iCol),Ecran);
if (Ecran.Name='GCARTDISPO') and not (ctxMode in V_PGI.PGIContexte) then
    begin
    SetControlVisible('TCOLLECTION',False) ;
    SetControlVisible('GA_COLLECTION',False) ;
    end ;
if (not VH_GC.GCMultiDepots) then
    SetControlProperty('RUPT1','Plus','AND CO_CODE<>"DEP"')
else SetControlProperty('RUPT1','Plus','AND CO_CODE<>"DEP"');
end ;

/////////CONTROLE DES RUPTURES PARAMETRABLES////////////////////////////////////////
Procedure TOF_DispoEtat_AffectGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value, St_Text, St : string;
    Indice, i_ind : integer;
BEGIN
F := TForm (Longint (Parms[0]));
if (copy(F.Name,1,10) <> 'GCARTDISPO') then exit;
Indice := Integer (Parms[1]);
if (not VH_GC.GCMultiDepots) then St_Plus := ' AND CO_CODE<>"DEP"'
else St_Plus := ' AND CO_CODE<>"ETA"';
St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value);
St_Text := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text);
For i_ind := 1 to 3 do
    BEGIN
    if i_ind = Indice then continue;
    St := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value);
    If St <> '' then St_Plus := St_Plus + ' AND CO_CODE <>"' + St + '"' ;
    END;
THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Plus := St_Plus;
THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value := St_Value;
THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text := St_Text ;
END;

Procedure TOF_DispoEtat_ChangeGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value : string;
    Indice, i_ind : integer;
BEGIN
F := TForm (Longint (Parms[0]));
if (copy(F.Name,1,10) <> 'GCARTDISPO') then exit;
Indice := Integer (Parms[1]);
St_Plus := THValComboBox (F.FindComponent ('RUPT'+InttoStr (Indice))).Plus;
St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value);
if St_Value = '' then
    BEGIN
    THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(Indice))).Text := '';
    THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text := '';
    THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Plus := '';
    THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text := '';
    For i_ind := Indice + 1 to 3 do
        BEGIN
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Enabled := True;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value := '';
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Plus := '';
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Enabled := False;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Color := clBtnFace;
        THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(i_ind))).Text := '';
        THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(i_ind))).Text := '';
        END;
    END else
    BEGIN
    if Indice < 3 then
        BEGIN
        THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice + 1))).Enabled := True;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice + 1))).Color := clWindow;
        END;
    THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(Indice))).Text :=
            RechDom ('GCGROUPDISPO', St_Value, True);
    THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text :=
            string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text);
    END;
END;

/////////////////////////////////////////////////////////////////////////////

procedure InitTOFDispoEtat ();
begin
RegisterAglProc('DispoEtat_ChangeGroup', True , 1, TOF_DispoEtat_ChangeGroup);
RegisterAglProc('DispoEtat_AffectGroup', True , 1, TOF_DispoEtat_AffectGroup);
end;

Initialization
registerclasses([Tof_DispoEtat]);
InitTOFDispoEtat();
end.
