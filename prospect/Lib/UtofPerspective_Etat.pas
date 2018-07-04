unit UtofPerspective_etat;

interface

uses  StdCtrls,Controls,Classes,forms,sysutils,
      HCtrls,HEnt1,UTOF, UtilGC,
{$ifdef AFFAIRE}
      UtofAfTraducChampLibre,
{$ENDIF}
{$IFDEF EAGLCLIENT}
     UTOB,
{$ELSE}
     db,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
        M3FP, Graphics,UtilSelection,UtilRT;

Type
{$ifdef AFFAIRE}
                //mcd 11/05/2006 12940  pour faire affectation depuis ressource si paramétré
     TOF_PerspEtat = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
    TOF_PerspEtat = Class (TOF)
{$endif}
      procedure OnLoad ; override ;
      procedure OnArgument(Arguments : String ) ; override ;
     END;

implementation
uses ParamSoc
{$ifdef GIGI}
  ,EntGC;
Var     Plus : string;
{$else}
;
{$endif}

procedure TOF_PerspEtat.OnArgument(Arguments : String ) ;
var i:integer;
begin
inherited ;
  MulCreerPagesCL(Ecran,'NOMFIC=GCTIERS');
  if (ecran <> Nil) and GetParamSocSecur('SO_RTGESTINFOS00V',False) then
    MulCreerPagesCL(Ecran,'NOMFIC=RTPERSPECTIVES');

for i:=1 to 3 do
    SetControlCaption('TRPE_TABLELIBREPER'+IntToStr(i),'&'+RechDom('RTLIBCHAMPSLIBRES','PL'+IntToStr(i),FALSE)) ;
  //mcd 22/05/07 le sligens ci-dessous était 2 fois... j'ai supprimé
if (GetControl('YTC_RESSOURCE1') <> nil)  then
  begin
  if not (ctxaffaire in V_PGI.PGICONTEXTE) then SetControlVisible ('PRESSOURCE',false)
  else begin
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '_');
    if not (ctxscot in V_PGI.PGICOntexte) then
       begin
       SetControlVisible ('T_MOISCLOTURE',false);
       SetControlVisible ('T_MOISCLOTURE_',false);
       SetControlVisible ('TT_MOISCLOTURE',false);
       SetControlVisible ('TT_MOISCLOTURE_',false);
       end;
    end;
  end;
{$Ifdef GIGI}
 SetControlText('TRPE_TYPETIERS',TraduireMemoire('Nature tiers'));
 SetControlText('T_NatureAuxi','');    //on efface les valeurs CLI et PO, car NCP en plus
 SetControlProperty ('T_NATUREAUXI', 'Complete', true);
 SetControlProperty ('T_NATUREAUXI', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('T_NATUREAUXI', 'Plus', VH_GC.AfNatTiersGRCGI);
 if (GetControl('RPE_REPRESENTANT') <> nil) then  SetControlVisible('RPE_REPRESENTANT',false);
 if (GetControl('TRPE_REPRESENTANT') <> nil) then  SetControlVisible('TRPE_REPRESENTANT',false);
 plus := ' AND CO_CODE <>"004"';
 if (GetControl('RPE_OPERATION') <> nil) and Not GetParamsocSecur('SO_AFRTOPERATIONS',False)
    then  begin
    SetControlVisible('RPE_OPERATION',false);
    SetControlVisible('TRPE_OPERATION',false);
    end;
 if (GetControl('RPE_PROJET') <> nil) and Not GetParamsocSecur('SO_RTPROJGESTION',False)
    then  begin
    SetControlVisible('RPE_PROJET',false);
    SetControlVisible('TRPE_PROJET',false);
    end;
{$endif}
{$IFDEF GRCLIGHT}
  if not GetParamsocSecur('SO_CRMACCOMPAGNEMENT',False) then
    begin
    SetControlVisible('RPE_OPERATION',false);
    SetControlVisible('TRPE_OPERATION',false);
    SetControlVisible('RPE_PROJET',false);
    SetControlVisible('TRPE_PROJET',false);
    end;
{$ENDIF GRCLIGHT}

end;

procedure TOF_PerspEtat.OnLoad;
Var F : TForm;
    xx_where,St : string;
    i : integer;
begin
Inherited;
F := TForm (Ecran);
{if (TCheckbox(F.FindComponent('PROPOPRINCIPALE')).Checked = true) then
   SetControlText ('XX_WHERE', 'RPE_VARIANTE="0" or RPE_PERSPECTIVE=RPE_VARIANTE')
else
   SetControlText ('XX_WHERE','');    }
xx_where := '';
if (F.Name = 'RTPERSP_ETAT') or (F.Name = 'RTPERSP_ETABLC') then
   if (TCheckbox(F.FindComponent('PROPOPRINCIPALE')).Checked = true) then
       xx_where := '(RPE_VARIANTE=0 or RPE_PERSPECTIVE=RPE_VARIANTE)';

xx_where := xx_where + RTXXWhereConfident('CON');
SetControlText('XX_WHERE',xx_where) ;

// nouveau critère de regroupement : date prévue, du coup, faut l'enlever du order by
//SetControlText('XX_ORDERBY', 'RPE_DATEREALISE,RPE_AUXILIAIRE');
St:='';
For i := 1 to 6 do
    BEGIN
    St := THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(i))).Text;
    If St = 'RPE_DATEREALISE' then
       break;
    END;
If St = 'RPE_DATEREALISE' then
   SetControlText('XX_ORDERBY', 'RPE_AUXILIAIRE')
else
   SetControlText('XX_ORDERBY', 'RPE_DATEREALISE,RPE_AUXILIAIRE');

end;

Procedure TOF_PerspEtat_AffectGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value, St_Text, St : string;
    Indice, i_ind : integer;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'RTPERSP_ETAT') and (F.Name <> 'RTPERSP_ETABLC') then exit;
Indice := Integer (Parms[1]);
St_Plus := '';
if not (ctxaffaire in V_PGI.PGIContexte) then St_Plus := 'AND CO_CODE not Like "R%"';
{$ifdef GIGI}
St_Plus := plus;
{$endif}
St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value);
St_Text := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text);
For i_ind := 1 to 6 do
    BEGIN
    if i_ind = Indice then continue;
    St := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value);
    If St <> '' then St_Plus := St_Plus + ' AND CO_CODE <>"' + St + '"';
    END;
THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Plus := St_Plus;
THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value := St_Value;
THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text := st_Text;
END;

Procedure TOF_PerspEtat_ChangeGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value : string;
    Indice, i_ind : integer;
    Q : TQuery;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'RTPERSP_ETAT') and (F.Name <> 'RTPERSP_ETABLC') then exit;
Indice := Integer (Parms[1]);
St_Plus := '';
{$ifdef GIGI}
St_Plus := plus;
{$endif}
St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value);
if St_Value = '' then
    BEGIN
    THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(Indice))).Text := '';
    THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text := '';
    TCheckBox(F.FindComponent('SAUTPAGE'+InttoStr(Indice))).Enabled := False;
    TCheckBox(F.FindComponent('SAUTPAGE'+InttoStr(Indice))).Checked := False;
    For i_ind := Indice + 1 to 6 do
        BEGIN
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Enabled := True;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value := '';
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Enabled := False;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Color := clBtnFace;
        THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(i_ind))).Text := '';
        THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(i_ind))).Text := '';
        TCheckBox(F.FindComponent('SAUTPAGE'+InttoStr(i_ind))).Enabled := False;
        TCheckBox(F.FindComponent('SAUTPAGE'+InttoStr(i_ind))).Checked := False;
        END;
    END else
    BEGIN
    if Indice < 6 then
        BEGIN
        THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice + 1))).Enabled := True;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice + 1))).Color := clWindow;
        END;
    TCheckBox(F.FindComponent('SAUTPAGE'+InttoStr(Indice))).Enabled := true;
    Q := OpenSQL('SELECT CO_LIBRE FROM COMMUN '+
         'WHERE CO_TYPE="RGS" AND CO_CODE="'+ St_Value+'"', true);

    THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(Indice))).Text := Q.Fields[0].AsString;
//  RechDom ('RTGROUPPERSP', St_Value, False);
    THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text :=
            string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text);
    Ferme(Q) ;
    END;
END;


Initialization
registerclasses([TOF_PerspEtat]);
RegisterAGLProc('Persp_ChangeGroup', True, 1, TOF_PerspEtat_ChangeGroup);
RegisterAGLProc('Persp_AffectGroup', True, 1, TOF_PerspEtat_AffectGroup);
end.
