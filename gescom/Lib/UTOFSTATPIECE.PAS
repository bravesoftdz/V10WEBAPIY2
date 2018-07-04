{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 15/12/2000
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : UTOFSTATPIECE ()
Mots clefs ... : TOF;UTOFSTATPIECE
*****************************************************************}
Unit UTOFSTATPIECE ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,  
     HCtrls, HEnt1, HMsgBox, UTOF, EntGC, Graphics, Vierge ;

Type
  TOF_GCSTATPIECE = Class (TOF)
    Private
    Nature,ResultatStat : String;
    CodePiece : array [1..3] of String;
    Public
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

Implementation

procedure TOF_GCSTATPIECE.OnUpdate ;
var i_ind : integer;
    Obligatoire : Boolean;
    TableLibre,Stat : String;
begin
  Inherited ;
  Stat:='';
  For i_ind := 1 to 3 Do
    begin
    if GetInfoParPiece(Nature,'GPP_PIECETABLE'+InttoStr(i_ind))<>'' then
      begin
      Obligatoire := GetInfoParPiece(Nature,'GPP_CODEPIECEOBL'+InttoStr(i_ind))='X';
      if Obligatoire then
         begin
         if THValComboBox(GetControl('GP_LIBREPIECE'+InttoStr(i_ind))).value='' then
           begin
           TableLibre:=THLabel(GetControl('TGP_PIECETABLE'+InttoStr(i_ind))).Caption;
           PGIBox('Vous devez renseigner la zone "'+ TableLibre+'" car elle est obligatoire',TFVierge(Ecran).Caption);
           SetFocusControl('GP_LIBREPIECE'+InttoStr(i_ind));
           TForm(Ecran).ModalResult:=0;
           exit;
           end;
         end;
      if Stat='' then Stat := InttoStr(i_ind)+'='+THValComboBox(GetControl('GP_LIBREPIECE'+InttoStr(i_ind))).value
      else Stat := Stat + ';'+InttoStr(i_ind)+'='+THValComboBox(GetControl('GP_LIBREPIECE'+InttoStr(i_ind))).value;
      end;
    end;
  ResultatStat := 'VALIDESTATPIECE=TRUE;'+Stat;
end ;

procedure TOF_GCSTATPIECE.OnLoad ;
var i_ind : integer;
    Plus,ValeurDefaut : string;
begin
  Inherited ;
  ResultatStat:='';
  For i_ind := 1 to 3 Do
    begin
    Plus := GetInfoParPiece(Nature,'GPP_PIECETABLE'+InttoStr(i_ind));
    if Plus<>'' then
      begin
      THLabel(GetControl('TGP_PIECETABLE'+InttoStr(i_ind))).Caption := RechDom('GCTYPESTATPIECE',Plus,False);
      THValComboBox(GetControl('GP_LIBREPIECE'+InttoStr(i_ind))).Plus := Plus;
      if CodePiece[i_ind]<>'' then THValComboBox(GetControl('GP_LIBREPIECE'+InttoStr(i_ind))).Value := CodePiece[i_ind]
      else
        begin
        ValeurDefaut := GetInfoParPiece(Nature,'GPP_CODPIECEDEF'+InttoStr(i_ind));
        if ValeurDefaut<>'' then THValComboBox(GetControl('GP_LIBREPIECE'+InttoStr(i_ind))).Value := ValeurDefaut;
        end;
      end
    else
      begin
      SetControlVisible('TGP_PIECETABLE'+InttoStr(i_ind),False);
      SetControlVisible('GP_LIBREPIECE'+InttoStr(i_ind),False);
      TForm(Ecran).Height := TForm(Ecran).Height - THValComboBox(GetControl('GP_LIBREPIECE'+InttoStr(i_ind))).Height;
      end;
    end;
end ;

procedure TOF_GCSTATPIECE.OnArgument (S : String ) ;
var Critere,ValMul,ChampMul : String;
    x,Indice : integer;
begin
  Inherited ;
Repeat
  Critere:=uppercase(Trim(ReadTokenSt(S))) ;
  if Critere<>'' then
    begin
    x:=pos('=',Critere);
    if x<>0 then
      begin
      ChampMul:=copy(Critere,1,x-1);
      ValMul:=copy(Critere,x+1,length(Critere));
      if ChampMul='NATUREPIECEG' then Nature:=ValMul
      else
        begin
        Indice:=StrToInt(copy(ChampMul,length(ChampMul),1));
        CodePiece[Indice]:=ValMul;
        end;
      end;
    end;
until  Critere='' ;
end ;

procedure TOF_GCSTATPIECE.OnClose ;
begin
  Inherited ;
  TFVierge(Ecran).FRetour := ResultatStat;
end ;

Initialization
  registerclasses ( [ TOF_GCSTATPIECE ] ) ;
end.
