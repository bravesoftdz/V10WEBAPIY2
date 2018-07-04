{***********UNITE*************************************************
Auteur  ...... : JT
Créé le ...... : 18/07/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : UTILULST ()
Suite ........ : Reçoie un TListBox et une TOB pour affichage du TListBox
Suite ........ : et éventuellement impression et export de la TOB
Mots clefs ... : TOF;UTILULST
*****************************************************************}
Unit UTILULST_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFNDEF EAGLCLIENT}
     db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} mul, FE_Main,
{$ELSE}
     MainEAGL, eMul,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, UTOB, HTB97,
     UtilUTilitaires, MailOl, Windows;

function GCLanceFiche_UtilULst(Nat,Cod : String ; Range,Lequel,Argument : string; TobData : TOB = nil; TListData : TStringList = nil; Titre : string = ''; FormEtat : integer = 0) : string;

Type
  TOF_UTILULST = Class (TOF)
    private
      procedure bImprimerClick(Sender : TObject);
      procedure bExporterClick(Sender : TObject);
      procedure bMailClick(Sender : TObject);
      procedure AfficheInfo(Titre, Lib : string);

    public
      procedure OnUpdate                 ; override ;
      procedure OnArgument (S : String ) ; override ;
      procedure OnClose                  ; override ;
  end ;

Implementation

var TobATraiter : TOB;
    LstATraiter : TListBox;
    TLstATraiter : TStringList;
    TitreEtat : string;
    FormatEtat : integer;

function GCLanceFiche_UtilULst(Nat,Cod : String ; Range,Lequel,Argument : string; TobData : TOB = nil; TListData : TStringList = nil; Titre : string = ''; FormEtat : integer = 0) : string;
begin
  Result := '';
  if Nat = '' then exit;
  if Cod = '' then exit;
  TobATraiter := TOB.Create('LES LIGNES',nil,-1);
  TLstATraiter := TStringList.Create;
  if TobData <> nil then
    TobATraiter.Dupliquer(TobData,True,True,True);
  if TListData <> nil then
    TLstATraiter := TListData;
  TitreEtat := Titre;
  FormatEtat := FormEtat;
  Result := AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_UTILULST.bImprimerClick(Sender : TObject);
var TobToPrint : TOB;
    Cpt : integer;
begin
  TobToPrint := TOB.Create('',nil,-1);
  for Cpt := 0 to LstATraiter.Items.Count -1 do
    UtilTobCreat(TobToPrint,'','',LstATraiter.Items[Cpt],'');
  UtilTobPrint(TobToPrint,TitreEtat,FormatEtat);
  TobToPrint.free;
end;

procedure TOF_UTILULST.bExporterClick(Sender : TObject);
begin
  UtilTobExportTxt(TobATraiter, TitreEtat);
end;

procedure TOF_UTILULST.bMailClick(Sender : TObject);
var OutLook : string;
    Txt : TStringList;
begin
  OutLook := GetFromRegistry(HKEY_LOCAL_MACHINE,'Software\Microsoft\Windows\CurrentVersion\App Paths\outlook.exe','Path',OutLook,TRUE) ;
  if OutLook <> '' then
  begin
    Txt := TStringList.Create;
    Txt.Text := LstATraiter.Items.Text;
    SendMail(TitreEtat,'','',HTStringList(Txt),'',False);
    Txt := nil;
    Txt.Free;
  end else
    PGiBox(TraduireMemoire('Envoie impossible, OutLook n''est pas installé sur votre ordinateur'),TraduireMemoire('Erreur'));
end;

procedure TOF_UTILULST.AfficheInfo(Titre, Lib : string);
var TobTmp : TOB;
    Sql : String;
    BlocNote : Variant;
    PosRC : integer;
begin
  TForm(Ecran).Caption := Titre;
  UpdateCaption(Ecran);
  TitreEtat := Titre;
  LstATraiter.Clear;
  Sql := 'SELECT * FROM JNALEVENT WHERE GEV_TYPEEVENT = "UTI" AND '+
         'GEV_LIBELLE LIKE "'+Lib+'" '+
         'ORDER BY GEV_DATEEVENT DESC, GEV_NUMEVENT DESC';
  TobTmp := TOB.Create('', nil, -1);
  TobTmp.LoadDetailFromSQL(Sql, True, False, 1);
  if TobTmp.Detail.Count > 0 then
  begin
    TOBTmp.Detail.Sort('GEV_DATEEVENT,GEV_NUMEVENT');
    LstATraiter.Items.Add(TraduireMemoire('Exécuté le : ')+
                          DateToStr(TobTmp.Detail[0].GetValue('GEV_DATEEVENT')));
    LstATraiter.Items.Add(TraduireMemoire('Par : ')+
                          RechDom('TTUTILISATEUR',TobTmp.Detail[0].GetValue('GEV_UTILISATEUR'),False));
    LstATraiter.Items.Add(TraduireMemoire('Description :'));
    BlocNote := TobTmp.Detail[0].GetValue('GEV_BLOCNOTE');
    while BlocNote <> '' do
    begin
      PosRC := pos(#13,BlocNote);
      LstATraiter.Items.Add(Copy(BlocNote,1,PosRC-1));
      BlocNote := Copy(BlocNote,PosRC+2,length(BlocNote));
    end;
  end else
  begin
    LstATraiter.Items.Add(TraduireMemoire('Cet utilitaire n''a jamais été exécuté.'));
  end;
  TobTmp.Free;
end;
{==============================================================================================}
{================================== Procédure de la TOF =======================================}
{==============================================================================================}
procedure TOF_UTILULST.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_UTILULST.OnArgument (S : String ) ;
var NumChoix : integer;
    Param : string;
begin
  Inherited ;
{ S = 0 : Vérification auxiliaire tiers
  S = 1 : RAZ activité
  S = 2 : Ajustement des compteurs
  S = 3 : Affectation dépôt
  S = 4 : Natures de doc sans modèle d'impression }
  Param := S;
  NumChoix := StrToInt(ReadTokenSt(Param));
  if Param <> '' then
    FormatEtat := StrToInt(ReadTokenSt(Param));
  LstATraiter := TListBox(GetControl('LSTRESULT'));
  TToolBarButton97(GetControl('BImprimer')).OnClick := bImprimerClick;
  TToolBarButton97(GetControl('BExporter')).OnClick := bExporterClick;
  TToolBarButton97(GetControl('BMail')).OnClick := bMailClick;
  if TLstATraiter.Text <> '' then
    LstATraiter.Items.Text := TLstATraiter.Text;
  //Lancement direct depuis le menu (sans assistant)
  if (NumChoix = 0) or (NumChoix = 4) then
  begin
    TForm(Ecran).Caption := TitreEtat;
    UpdateCaption(Ecran);
  end else
  begin
    TToolBarButton97(GetControl('BExporter')).Visible := False;
    case NumChoix of
        1 : AfficheInfo(TraduireMemoire('RAZ de l''activité'), TraduireMemoire('Remise à zéro de%'));
        2 : AfficheInfo(TraduireMemoire('Recalcul compteur'), TraduireMemoire('Ajustement des compteurs%'));
        3 : AfficheInfo(TraduireMemoire('Affectation dépôt'), TraduireMemoire('Affectation dépôt%'));
    end;
  end;
end ;

procedure TOF_UTILULST.OnClose ;
begin
  Inherited ;
  if TLstATraiter <> nil then
  begin
    TLstATraiter := nil;
    TLstATraiter.free;
  end;
  if TobATraiter <> nil then
  begin
    TobATraiter := nil;
    TobATraiter.free;
  end;
end ;

Initialization
  registerclasses ( [ TOF_UTILULST ] ) ;
end.

