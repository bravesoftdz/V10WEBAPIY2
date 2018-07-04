{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 18/01/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MBOVALIDEMODELE ()
Mots clefs ... : TOF;MBOVALIDEMODELE
*****************************************************************}
Unit UTOFMBOVALIDEMODELE ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,Vierge,Spin,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF ; 

Type
  TOF_MBOVALIDEMODELE = Class (TOF)

  public
         Action     : TActionFiche ;
         NatureEtat : String;
         modeleEtat : String;
         Apercu     : String;
         ApercuEtiq : String;
         ModeleEtiq : String;
         ImpEtiq    : String ;
         nbCopie    : Integer ;
         Sav        : Boolean ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (Argument : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

Implementation

procedure TOF_MBOVALIDEMODELE.OnUpdate ;
var F: TFVierge ;
    BApercu : TCheckBox ;
begin
  Inherited ;
  if not (Ecran is TFVierge) then exit;
  F:=TFVierge(Ecran) ;

  ModeleEtat := THValComboBox((F).FindComponent('IMPETAT')).Value ;
  BApercu    := TCheckBox((F).FindComponent('APERCUAVIMP')) ;
  NbCopie    := TSpinEdit((F).FindComponent('NBEXEMPLAIRE')).Value ;

  // Etiquettes
  ModeleEtiq:=THValComboBox((F).FindComponent('ETATETIQ')).Value ;

  if Bapercu.Checked=True then
    Apercu:='X'
  else
    Apercu:='-' ;

  if TCheckBox((F).FindComponent('APERCUAVETIQ')).Checked=True then
    ApercuEtiq:='X'
  else
    ApercuEtiq:='-' ;

  if TCheckBox((F).FindComponent('IMPETIQ')).Checked=True      then
    ImpEtiq:='X'
  else
    ImpEtiq:='-' ;

  //if ModeleEtat<>'' then
  F.Retour:='X;'+ModeleEtat+';'+Apercu+';'+IntToStr(NbCopie)+';'+ModeleEtiq+';'+ApercuEtiq+';'+ImpEtiq+'' ;//else
  //F.Retour:='-;'+ModeleDoc+';'+Apercu+';'+IntToStr(NbCopie)+';'+ModeleEtiq+';'+ApercuEtiq+';'+ImpEtiq+'';

  Sav:=True ;

end ;

procedure TOF_MBOVALIDEMODELE.OnLoad ;
Var F: TFVierge ;
BApercu: TCheckBox ;
begin
  Inherited ;

  if not (Ecran is TFVierge) then exit;

  F:=TFVierge(Ecran) ;

  // Pour la mode il faut le bon modele d'étiquette trouvé par la tablette GCIMPETIQMODE
  if (ctxMode in V_PGI.PGIContexte) then THValComboBox(F.FindComponent ('ETATETIQ')).DataType:='GCIMPETIQMODE';

  //FV1 : 09/06/2017 - FS#2585 - VEODIS GROUP : Bon de Commande N° GB-170515-1
  if NatureEtat <> '' then
  begin
    THValComboBox((F).FindComponent('IMPETAT')).DataType := 'BTMODELETAT';
    THValComboBox((F).FindComponent('IMPETAT')).Plus     := 'MO_NATURE="' + NatureEtat + '"';
  end;
  THValComboBox((F).FindComponent('IMPETAT')).Value := ModeleEtat;

  THValComboBox((F).FindComponent('ETATETIQ')).Value:= ModeleEtiq;
  BApercu:=TCheckBox((F).FindComponent('APERCUAVIMP')) ;

  if Apercu='X' then Bapercu.state:=cbChecked else Bapercu.state:=cbUnChecked ;

  TSpinEdit((F).FindComponent('NBEXEMPLAIRE')).Value:=NbCopie ;

  // Etiquettes
  if ApercuEtiq='X' then
    TCheckBox((F).FindComponent('APERCUAVETIQ')).State:=cbChecked
  else
    TCheckBox((F).FindComponent('APERCUAVETIQ')).State:=cbUnChecked ;

  if ImpEtiq='X' then
    TCheckBox((F).FindComponent('IMPETIQ')).State:=cbChecked
  else
    TCheckBox((F).FindComponent('IMPETIQ')).State:=cbUnChecked ;

  TGroupBox((F).FindComponent('EDPIECE')).SetFocus ;

  Sav:=False ;

end ;

procedure TOF_MBOVALIDEMODELE.OnArgument (Argument : String ) ;
Var St,ChampMul,ValMul,Critere,stTitre:String ;
X: Integer ;
begin
  Inherited ;

  St    :=Argument ;
  Action:=taModif ;

  Repeat
    Critere:=Trim(ReadTokenSt(Argument)) ;
    if Critere<>'' then
    begin
      x:=pos('=',Critere);
      if x<>0 then
      begin
        ChampMul:=copy(Critere,1,x-1) ;
        ValMul:=copy(Critere,x+1,length(Critere)) ;
        if ChampMul='NATUREETAT' then NatureEtat:=ValMul ;
        if ChampMul='MODELEETAT' then ModeleEtat:=ValMul ;
        if ChampMul='MODELEETIQ' then ModeleEtiq:=ValMul ;
        if ChampMul='APERCU'     then Apercu:=ValMul ;
        if ChampMul='APERCUETIQ' then ApercuEtiq:=ValMul ;
        if ChampMul='IMPETIQ'    then ImpEtiq:=ValMul ;
        if ChampMul='NBCOPIE'    then NbCopie:=StrToInt(ValMul) ;
        if ChampMul='TITRE'      then stTitre:=ValMul ;
      end ;
    end ;
  until  Argument='' ;

  if sttitre <> '' then ecran.caption := stTitre;

end ;

procedure TOF_MBOVALIDEMODELE.OnClose ;
Var F: TFVierge ;
begin
  Inherited ;
  if not (Ecran is TFVierge) then exit;
  F:=TFVierge(Ecran) ;
  if not Sav then F.Retour:='' ;
end ;

Initialization
  registerclasses ( [ TOF_MBOVALIDEMODELE ] ) ; 
end.
