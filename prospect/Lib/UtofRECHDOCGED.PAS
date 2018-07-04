{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 06/12/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : RECHDOCGED ()
Mots clefs ... : TOF;RECHDOCGED
*****************************************************************}
Unit UtofRECHDOCGED ;

Interface

Uses 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     mul,
     Fe_Main,
{$else}
     eMul, 
     MaineAGL,
{$ENDIF}
     forms, 
     sysutils, 
     HCtrls,
     HEnt1, 
     HMsgBox, 
     UTOF,
     HTB97,
     RTNewDocument,
     UGedFileViewer,
     UtilGEDRT,
     UtilGC,
     UtilRT
 ;

Type
  TOF_RECHDOCGED = Class (TOF)
  private
    Reference1,Objet,xx_where : string;
    Reference2 : integer;
    procedure SupprimeDocument;
  Public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure FListeOnDblClick(Sender: TObject);
    procedure  BtOnClick(Sender: TObject);
    procedure  DispatchGedGRC(NumTag: Integer);
  end ;

Const
  // Tags
  cTagSupprimer           = 5; // Supprimer le document
  cTagProprietes          = 4; // Propriétés du document
  cTagDocExterne          = 1; // Insertion de document externe
  cTagApercu              = 2; // Visualiser le document

Function RTLanceFiche_RECHDOCGED(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation

{$IFDEF VER150}
  uses
    Variants
    ;
{$ENDIF VER150}

procedure TOF_RECHDOCGED.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_RECHDOCGED.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_RECHDOCGED.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_RECHDOCGED.OnLoad ;
begin
  Inherited ;
  if objet = '' then SetControlText('XX_WHERE',xx_where) ;
end ;

procedure TOF_RECHDOCGED.OnArgument (S : String ) ;
var  Critere,ChampMul,ValMul : string;
     x : integer;
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
           if ChampMul='TIERS' then
           begin
              Reference1 := ValMul;
           end
           else if (ChampMul='CHAINAGE') or (ChampMul='ACTION') then
                begin
                Reference2 := Valeuri(ValMul);
                end
           else if ChampMul='OBJET' then
                begin
                Objet := ValMul;
                end;
        end;
    end;
  until  Critere='';
  MajChampsLibresGED(TForm(Longint(ecran)));
  xx_where:=RTXXWhereConfident('CON');
  THGrid(GetControl('FLISTE')).OnDblClick := FListeOnDblClick;
  TToolbarButton97(GetControl('BPROPRIETE')).OnClick := BtOnClick;
  TToolbarButton97(GetControl('BDELETE')).OnClick := BtOnClick;
  TToolbarButton97(GetControl('BINSERT')).OnClick := BtOnClick;
  if (objet = 'CHA') or (objet = 'ACT') or (objet = 'PRO') then SetControlVisible ('PCOMPLEMENT',False);
end ;

procedure TOF_RECHDOCGED.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_RECHDOCGED.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_RECHDOCGED.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_RECHDOCGED.FListeOnDblClick(Sender: TObject);
var sTempFileName, sTitre: String;
begin
  inherited;
  if Not VarIsNull(GetField('RTD_DOCGUID')) then
    begin
    sTempFileName := ExtraitDocumentRT(GetField('RTD_DOCGUID'), sTitre);
    if sTempFileName='' then exit;
    ShowGedFileViewer(sTempFileName, False, sTitre, True, False, True, True);
    end;
end;

procedure TOF_RECHDOCGED.BtOnClick(Sender: TObject);
var Num : Integer;
begin
  Num := 0;
  if Sender is TToolBarButton97 then
    Num := TToolBarButton97(Sender).Tag;
  DispatchGedGRC(Num);
end;

procedure TOF_RECHDOCGED.DispatchGedGRC(NumTag: Integer);
var St, Infos : String;
    ParGed: TParamGedDoc;
begin
  case NumTag of

  // Supprimer
  cTagSupprimer :
      begin
      if Not VarIsNull(GetField('RTD_DOCGUID')) then
        begin
        SupprimeDocument;
        TFMul(Ecran).BChercheClick(Nil);
        end;
      end;

  // Propriétés
  cTagProprietes :
      begin
      if Not VarIsNull(GetField('RTD_DOCGUID')) then
        begin
        ParGed.SDocGUID := GetField('RTD_DOCGUID');
        ParGed.SFileGUID := GetField('YDF_FileGUId');
        ParGed.NoDossier := '';
        ParGed.CodeGed := '';
        ParGed.TypeGed := '' ;
        ParGed.Objet := Objet;
        ParGed.ModifLien := True;
        if Objet = 'TIE' then Infos := 'Tiers='+GetField('RTD_TIERS')+';'
        else if Objet = 'CHA' then Infos := 'Tiers='+GetField('RTD_TIERS')+';CHAINAGE='+InttoStr(GetField('RTD_NUMCHAINAGE'))+';'
        else if Objet = 'ACT' then Infos := 'Tiers='+GetField('RTD_TIERS')+';ACTION='+InttoStr(GetField('RTD_NUMACTION'))+';'
        else if Objet = 'PRO' then Infos := 'Tiers='+GetField('RTD_TIERS')+';PROPOSITION='+InttoStr(GetField('RTD_PERSPECTIVE'))+';';
        Parged.Infos := Infos;
        St := ShowNewDocument(ParGed);
        TFMul(Ecran).BChercheClick(Nil);
        end;
      end;

  // Insertion de document externe
  cTagDocExterne :
      begin
      ParGed.SDocGUID := '';
      ParGed.SFileGUID := '';
      ParGed.NoDossier := '';
      ParGed.CodeGed := '';
      ParGed.Objet := Objet;
      if Objet = '' then ParGed.ModifLien := True
      else ParGed.ModifLien := False;
      if Objet = 'TIE' then Infos := 'Tiers='+GetControlText('RTD_TIERS')+';'
      else if Objet = 'CHA' then Infos := 'Tiers='+GetControlText('RTD_TIERS')+';CHAINAGE='+GetControlText('RTD_NUMCHAINAGE')+';'
      else if Objet = 'ACT' then Infos := 'Tiers='+GetControlText('RTD_TIERS')+';ACTION='+GetControlText('RTD_NUMACTION')+';'
      else if Objet = 'PRO' then Infos := 'Tiers='+GetControlText('RTD_TIERS')+';PROPOSITION='+GetControlText('RTD_PERSPECTIVE')+';';
      Parged.Infos := Infos;
      St := ShowNewDocument(ParGed);
      TFMul(Ecran).BChercheClick(Nil);
//      TFMul(Ecran).FListe.Refresh;
      end;

  end;

end;

procedure TOF_RECHDOCGED.SupprimeDocument;
var Msg : String;
begin
{    Msg := 'Vous allez supprimer la référence au document '+GetField('YDO_LIBELLEDOC')+'.'+#13#10
    + ' Rq : si le document est référencé par ailleurs, il ne sera pas supprimé de la GED.'+#13#10
    + ' Confirmez-vous la suppression ?';        }
    Msg := 'Suppression du document '+GetField('YDO_LIBELLEDOC')+'.'+#13#10
    + ' Confirmez-vous la suppression ?';

  if PGIAsk(Msg, TitreHalley)=mrNo then exit;

  // Suppression en cascade dans les tables, avec tests de dépendance
  SupprimeDocumentGedRT(GetField('RTD_DOCGUID'));
end;

Function RTLanceFiche_RECHDOCGED(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

Initialization
  registerclasses ( [ TOF_RECHDOCGED ] ) ; 
end.
