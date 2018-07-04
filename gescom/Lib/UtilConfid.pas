unit UtilConfid;

interface

uses  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
{$IFDEF EAGLCLIENT}
      UtileAGL,eFiche,eFichList,maineagl,Vierge,
{$ELSE}
      DBCtrls,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FichList,Fiche,Fe_Main,Vierge,
{$IFDEF V530}
      EdtEtat,
{$ELSE}
      EdtREtat,
{$ENDIF}
      MAJTable,HDB,
{$ENDIF}                                                                  // mng
      StdCtrls, ExtCtrls, Hctrls, UTOB, UTOM, UTOF, Grids, hmsgbox, HEnt1,Comctrls,SaisieList;

function  VerifierChampsObligatoires(F : TForm; Abrege : string; EcranName : string='';ChDataSet : Boolean=True) : string;
procedure AppliquerConfidentialite(F : TForm; Abrege : string);

implementation
uses   CbpMCD,
  		 CbpEnumerator;

function VerifierChampsObligatoires(F : TForm; Abrege : string; EcranName : string=''; ChDataSet : Boolean=True) : string;
var  OM : TOM ;
     ind1 : integer;
     CodeFiche, NomChamp, NomTable : string;
     TobFic, TobTemp : TOB;
     TCompTmp : TComponent;
     Q : TQuery;
     MaTOF : TOF ;
     BTOM,BTOF : Boolean;
		 Mcd : IMCDServiceCOM;
begin

Result := ''; BTOM:=False; BTOF:=False;
OM:=Nil; MaTof:=Nil;
if (F is TFFiche) then begin BTOM:=True; OM := TFFiche(F).OM ; end else
   if (F is TFFicheListe) then begin BTOM:=True; OM := TFFicheListe(F).OM ; end else
      if (F is TFVierge) then begin BTOF:=True; MaTOF := TFVierge(F).LaTOF ; end else Exit;
if EcranName = '' then
NomChamp := 'Select * from COMMUN where CO_TYPE="POB" and CO_ABREGE="' + Abrege + '" and ' +
            'CO_LIBRE like "%' + F.Name + '%"'
else
NomChamp := 'Select * from COMMUN where CO_TYPE="POB" and CO_ABREGE="' + Abrege + '" and ' +
            'CO_LIBRE like "%' + EcranName + '%"' ;
            
//modif gm 06/10/03
if (ctxaffaire in V_PGI.PGIContexte)  then
   NomChamp :=NomChamp +' AND ( CO_CODE like "A%" OR CO_CODE like "B%" '
else
  if (ctxGCAFF in V_PGI.PGIContexte)  then
   NomChamp :=NomChamp +' AND ( CO_CODE like "G%" or CO_CODE like "A%" '
  else
    NomChamp :=NomChamp +' AND ( CO_CODE like "G%" ';

if ctxGRC in V_PGI.PGIContexte then NomChamp :=NomChamp + ' OR CO_CODE like "R%")'
else  NomChamp :=NomChamp + ')';

{$IFDEF CHR}
NomChamp := 'Select * from COMMUN where CO_TYPE="POB" and CO_ABREGE="' + Abrege + '" and ' +
            'CO_LIBRE like "%' + F.Name + '%"';
{$ENDIF}

Q := OpenSQL(NomChamp, True,-1,'COMMUN');
if Q.EOF then
    begin
    Ferme(Q);
    Exit;
    end;
TobFic := TOB.Create('', nil, -1);
TobFic.LoadDetailDB('COMMUN', '', '', Q, False);
Ferme(Q);
CodeFiche := TobFic.Detail[0].GetValue('CO_CODE');
NomTable  := TobFic.Detail[0].GetValue('CO_LIBRE');
Delete(NomTable, Pos(';', NomTable), 255);
TobTemp := TOB.Create(NomTable, nil, -1);
TobFic.Free;

NomChamp := 'Select GOB_NOMCHAMP from PARAMOBLIG where GOB_CODE="' + CodeFiche + '" and ' +
            'GOB_OBLIGATOIRE="X"';
Q := OpenSQL(NomChamp, True,-1,'PARAMOBLIG');
if Q.EOF then
    begin
    Ferme(Q);
    TobTemp.Free;
    Exit;
    end;
TobFic := TOB.Create('', nil, -1);
TobFic.LoadDetailDB('PARAMOBLIG', '', '', Q, False);
Ferme(Q);
Result := '';
for ind1 := 0 to F.ComponentCount - 1 do
    begin
    TCompTmp := F.Components[ind1];
    NomChamp := TCompTmp.Name;

    if (TobFic.FindFirst(['GOB_NOMCHAMP'], [NomChamp], True) = nil) then Continue;
    if BTOM=True then
       begin
       if not (OM.GetControlVisible(NomChamp)) then Continue;
       if not (OM.GetControlEnabled(NomChamp)) then Continue;
       end;
    if BTOF=True then
       begin
       if not (MATOF.GetControlVisible(NomChamp)) then Continue;
       if not (MATOF.GetControlEnabled(NomChamp)) then Continue;
       end;

{$IFNDEF EAGLCLIENT}
    if (TCompTmp is THDBEdit) and THDBEdit(TCompTmp).ReadOnly then Continue;
    if (TCompTmp is THDBValComboBox) and THDBValComboBox(TCompTmp).ReadOnly then Continue;
{$ENDIF}
    if (TCompTmp is THEdit) and THEdit(TCompTmp).ReadOnly then Continue;
    if (TCompTmp is TEdit) and TEdit(TCompTmp).ReadOnly then Continue;

    if (Mcd.GetField(NomChamp).tipe = 'DOUBLE') then
        begin
        if BTOM=True then
           begin
           if ChDataSet then
             begin
             if (OM.GetField(NomChamp) = 0.0) then Result := Result + NomChamp + ';' ;
             end else if (StrToFloat(OM.GetControlText(NomChamp)) = 0.0) then Result := Result + NomChamp + ';';
           end else 
           begin
           if (StrToFloat(MATOF.GetControlText(NomChamp)) = 0.0) then Result := Result + NomChamp + ';';
           end;
        end
        else if (Mcd.GetField(NomChamp).tipe = 'DATE') then
        begin
        if BTOM=True then
           begin
           if ChDataSet then
             begin
             if (OM.GetField(NomChamp) = iDate1900) then Result := Result + NomChamp + ';' ;
             end else if (OM.GetControlText(NomChamp)='') or (OM.GetControlText(NomChamp)=DateToStr(iDate1900)) then Result := Result + NomChamp + ';' ;
           end else
           begin
           if (MATOF.GetControlText(NomChamp)='') or (MATOF.GetControlText(NomChamp)=DateToStr(iDate1900)) then Result := Result + NomChamp + ';';
           end;
        end
        else
        if BTOM=True then
           begin
           if ChDataSet then
             begin
             if (OM.GetField(NomChamp) = '') then Result := Result + NomChamp + ';' ;
             end else if (OM.GetControlText(NomChamp) = '') then Result := Result + NomChamp + ';' ;
           end else
           begin
           if (MATOF.GetControlText(NomChamp) = '') then Result := Result + NomChamp + ';';
           end;
    end;
TobFic.Free;
TobTemp.Free;
end;

procedure AppliquerConfidentialite(F : TForm; Abrege : string);
var OM : TOM ;
    ind1, Option : integer;       // mng
    CodeFiche, NomChamp, NomTable,Propriete : string;
    TobFic, TobTemp : TOB;
    TCompTmp : TComponent;
    Q : TQuery;
    bVisibleOnly : boolean;
    MaTOF : TOF ;
    BTOM{,BTOF }: Boolean;

begin
OM := Nil; MaTOF :=Nil;
BTOM := false;
                  // mng
if (F is TFFiche) then begin BTOM:=True; OM := TFFiche(F).OM ; end else
   if (F is TFFicheListe) then begin BTOM:=True; OM := TFFicheListe(F).OM ; end else
      if (F is TFSaisieList) then begin BTOM:=True; OM:=TFSaisieList(F).LeFiltre.OM ; end else
         if (F is TFVierge) then begin {BTOF:=True;} MaTOF := TFVierge(F).LaTOF ; end else Exit;
NomChamp := 'Select * from COMMUN where CO_TYPE="POB" and CO_ABREGE="' + Abrege + '" and ' +
            'CO_LIBRE like "%' + F.Name + '%"';

            //modif gm 06/10/03
if (ctxaffaire in V_PGI.PGIContexte)  then
   NomChamp :=NomChamp +' AND ( CO_CODE like "A%" OR CO_CODE like "B%" '
else
  if (ctxGCAFF in V_PGI.PGIContexte)  then
   NomChamp :=NomChamp +' AND ( CO_CODE like "G%" or CO_CODE like "A%" '
  else
    NomChamp :=NomChamp +' AND ( CO_CODE like "G%" ';

if ctxGRC in V_PGI.PGIContexte then NomChamp :=NomChamp +' or CO_CODE like "R%" )'
else NomChamp :=NomChamp +')';
{$IFDEF CHR}
NomChamp := 'Select * from COMMUN where CO_TYPE="POB" and CO_ABREGE="' + Abrege + '" and ' +
            'CO_LIBRE like "%' + F.Name + '%"';
{$ENDIF}
Q := OpenSQL(NomChamp, True,-1,'COMMUN');
if Q.EOF then
    begin
    Ferme(Q);
    Exit;
    end;
TobFic := TOB.Create('', nil, -1);
TobFic.LoadDetailDB('COMMUN', '', '', Q, False);
Ferme(Q);
CodeFiche := TobFic.Detail[0].GetValue('CO_CODE');
NomChamp  := TobFic.Detail[0].GetValue('CO_LIBRE');
Delete(NomChamp, 1, Pos(';', NomChamp));
NomTable := ReadTokenSt(NomChamp);
if NomChamp<>'' then Option:=StrToInt(NomChamp) else Option:=1;
TobFic.Free;

NomChamp := 'Select * from PARAMOBLIG where GOB_CODE="' + CodeFiche + '" and ' +
            '(GOB_GRVISIBLE<>"" or GOB_GRENABLE<>"" or GOB_USVISIBLE1<>"" or GOB_USENABLE1<>"")';
Q := OpenSQL(NomChamp, True,-1,'PARAMOBLIG');
if Q.EOF then
    begin
    Ferme(Q);
    Exit;
    end;
TobFic := TOB.Create('', nil, -1);
TobFic.LoadDetailDB('PARAMOBLIG', '', '', Q, False);
Ferme(Q);

for ind1 := 0 to TobFic.Detail.Count - 1 do
    begin
    TobTemp := TobFic.Detail[ind1];
    if ind1=0 then Option:=StrToInt(TobTemp.GetValue('GOB_PRIOCONFID'));
    TobTemp.AddChampSup('GOB_USVISIBLE', True);
    TobTemp.AddChampSup('GOB_USENABLE', True);
    TobTemp.PutValue('GOB_USVISIBLE', Trim( TobTemp.GetValue('GOB_USVISIBLE1') +
                                            TobTemp.GetValue('GOB_USVISIBLE2')));
    TobTemp.PutValue('GOB_USENABLE', Trim( TobTemp.GetValue('GOB_USENABLE1') +
                                           TobTemp.GetValue('GOB_USENABLE2')));
    end;

for ind1 := 0 to F.ComponentCount - 1 do
    begin
    bVisibleOnly := False;
    TCompTmp := F.Components[ind1];
    NomChamp := TCompTmp.Name;
    // mng
    if TCompTmp.ClassType = TTabSheet then Propriete:='TabVisible'
    else Propriete:='Visible';
    TobTemp := TobFic.FindFirst(['GOB_NOMCHAMP'], [NomChamp], True);
    if TobTemp = nil then
        begin
        if (TCompTmp is THLabel) or (TCompTmp is TLabel) then
            begin
            if TLabel(TCompTmp).FocusControl = nil then Continue;
            bVisibleOnly := True;
            TobTemp := TobFic.FindFirst(['GOB_NOMCHAMP'], [TLabel(TCompTmp).FocusControl.Name], True);
            if TobTemp = nil then Continue;
            end
            else Continue;
        end;
    case Option of
//  le groupe est prioritaire sur l'utilisateur
    1 : begin
        if Pos(V_PGI.Groupe, TobTemp.GetValue('GOB_GRVISIBLE')) <> 0 then
           if BTOM=True then OM.SetControlProperty(NomChamp, Propriete, False)
                        else MATOF.SetControlProperty(NomChamp, Propriete, False);
        if not bVisibleOnly then
            if Pos(V_PGI.Groupe, TobTemp.GetValue('GOB_GRENABLE')) <> 0 then
               if BTOM=True then OM.SetControlProperty(NomChamp, 'Enabled', False)
                            else MATOF.SetControlProperty(NomChamp, 'Enabled', False);
        end;
//  l'utilisateur est prioritaire sur le groupe
    2 : begin
        if Pos(V_PGI.User, TobTemp.GetValue('GOB_USVISIBLE')) <> 0 then
           if BTOM=True then OM.SetControlProperty(NomChamp, Propriete,False)
                        else MATOF.SetControlProperty(NomChamp, Propriete,False);
        if not bVisibleOnly then
            if Pos(V_PGI.User, TobTemp.GetValue('GOB_USENABLE')) <> 0 then
               if BTOM=True then OM.SetControlProperty(NomChamp, 'Enabled', False)
                            else MATOF.SetControlProperty(NomChamp, 'Enabled', False);
        end;
//  VRAI si l'utilisateur ET le groupe presents
    3 : begin
        if (Pos(V_PGI.Groupe, TobTemp.GetValue('GOB_GRVISIBLE')) <> 0) and
           (Pos(V_PGI.User, TobTemp.GetValue('GOB_USVISIBLE')) <> 0) then
           if BTOM=True then OM.SetControlProperty(NomChamp, Propriete,False)
                        else MATOF.SetControlProperty(NomChamp, Propriete,False);
        if not bVisibleOnly then
            if (Pos(V_PGI.Groupe, TobTemp.GetValue('GOB_GRENABLE')) <> 0) and
               (Pos(V_PGI.User, TobTemp.GetValue('GOB_USENABLE')) <> 0) then
               if BTOM=True then OM.SetControlProperty(NomChamp, 'Enabled', False)
                            else MATOF.SetControlProperty(NomChamp, 'Enabled', False);
        end;
//  VRAI si l'utilisateur OU le groupe presents
    4 : begin
        if (Pos(V_PGI.Groupe, TobTemp.GetValue('GOB_GRVISIBLE')) <> 0) or
           (Pos(V_PGI.User, TobTemp.GetValue('GOB_USVISIBLE')) <> 0) then
           if BTOM=True then OM.SetControlProperty(NomChamp, Propriete,False)
                        else MATOF.SetControlProperty(NomChamp, Propriete,False);
        if not bVisibleOnly then
            if (Pos(V_PGI.Groupe, TobTemp.GetValue('GOB_GRENABLE')) <> 0) or
               (Pos(V_PGI.User, TobTemp.GetValue('GOB_USENABLE')) <> 0) then
               if BTOM=True then OM.SetControlProperty(NomChamp, 'Enabled', False)
                            else MATOF.SetControlProperty(NomChamp, 'Enabled', False);
        end;
    end;
    end;
TobFic.Free;
end;

end.
