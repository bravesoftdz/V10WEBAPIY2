unit YMessageDetailDestinataire;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, StdCtrls, Hctrls, ComCtrls, HRichEdt, HRichOLE, HSysMenu, HMsgBox,
  HTB97, ExtCtrls, utob
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} ,dbtables {$ELSE} ,uDbxDataSet {$ENDIF}
{$ENDIF EAGLCLIENT}
  ;

type
  TFMsgDetailDest = class(TFVierge)
    Panel1: TPanel;
    HLabel3: THLabel;
    HRDestinataireA: THRichEditOLE;
    HLabel2: THLabel;
    HRDestinataireCc: THRichEditOLE;
    procedure FormShow(Sender: TObject);

  private
    MsgGuid: string;
    TobDestA: Tob;
    TobDestCc: Tob;
  public
  end;

//------------------------------------
//--- Déclaration procedure globale
//------------------------------------
procedure AfficherDetailDestinataire(UnMsgGuid: string; TobDestA: Tob; TobDestCc: Tob);
procedure InitialiserTobListeDestinataire (UneTob : Tob; ListeDestinataire : String; MsgGuid : String='');

function DonnerListeDestinataireFromSql (UnMsgGuid: string; TypeDest : String; ListeDestinataire: string) : String;
function DonnerListeDestinataireFromTob (UneTob : Tob; Detail : Boolean) : String;

function DonnerLibelleDestinataire (MsgGuid : String; UnDestinataire: string) : String;
function DonnerElementListeDistribution (UneListeDistribution : String) : String;

procedure SupprimerDestinataireDansTob (UneTob : Tob; UneListeDestinataire : String);
procedure AjouterDestinataireDansTob (UneTob : Tob; UneListeDestinataire : String);

implementation

uses YRechercherDestinataire;
{$R *.DFM}

//---------------------------------------------
//--- Nom   : RechercherDestinataire
//--- Objet :
//---------------------------------------------
procedure AfficherDetailDestinataire(UnMsgGuid: string; TobDestA: Tob; TobDestCc: Tob);
var FMsgDetailDest: TFMsgDetailDest;
begin
  FMsgDetailDest := TFMsgDetailDest.Create(Application);
  FMsgDetailDest.MsgGuid := UnMsgGuid;
  FMsgDetailDest.TobDestA := TobDestA;
  FMsgDetailDest.TobDestCc := TobDestCc;

  try
    FMsgDetailDest.ShowModal;
  finally
    FMsgDetailDest.Free;
  end;
end;

//---------------------------------------------
//--- Nom   : FormShow
//--- Objet :
//---------------------------------------------

procedure TFMsgDetailDest.FormShow(Sender: TObject);
var UneListeDestinataireA, UneListeDestinataireCc : String;
begin
  inherited;
  UneListeDestinataireA:=DonnerListeDestinataireFromTob (TobDestA,True);
  UneListeDestinataireCc:=DonnerListeDestinataireFromTob (TobDestCc,True);
  StringToRich(HRDestinataireA, UneListeDestinataireA);
  StringToRich(HRDestinataireCc, UneListeDestinataireCc);
end;

//--------------------------------------------
//--- Nom : InitialiserTobListeDestinataire
//--------------------------------------------
procedure InitialiserTobListeDestinataire (UneTob : Tob; ListeDestinataire : String; MsgGuid : String='');
var UnDestinataire      : String;
    UnLibelle           : String;
    UnType              : String;
    UneInfoDestinataire : String;
    UneTobEnreg         : Tob;
begin
 if (UneTob<>nil) then
  begin
   while (ListeDestinataire<>'') do
    begin
     UneInfoDestinataire:=ReadToKenPipe (ListeDestinataire,';');
     UnDestinataire:=ReadTokenPipe (UneInfoDestinataire,'|');
     UnLibelle:=ReadToKenPipe (UneInfoDestinataire,'|');
     UnType:=ReadToKenPipe (UneInfoDestinataire,'|');
     if (UneTob.FindFirst (['Nom'],[UnDestinataire],True)=nil) then
      begin
       if (UnLibelle='') then UnLibelle:=DonnerLibelleDestinataire (MsgGuid,UnDestinataire);
       UneInfoDestinataire:=UnDestinataire+'|'+UnLibelle+'|'+UnType;
       UneTobEnreg:=Tob.Create ('',UneTob,-1);
       UneTobEnreg.LoadFromSt ('Nom|Libelle|Type','|',UneInfoDestinataire,'|');
      end
    end;
  end;
end;

//-----------------------------------------
//--- Nom : DonnerLibelleDestinataire
//-----------------------------------------
function DonnerLibelleDestinataire (MsgGuid : String; UnDestinataire: string) : String;
var SSql                   : String;
    RSql                   : TQuery;
    UnLibelle              : String;
begin
 UnLibelle:='';

 //--- Recherche des libelles dans la table YMSGADDRESS
 if (MsgGuid<>'') then
  begin
   SSql := 'SELECT YMR_EMAILNAME FROM YMSGADDRESS WHERE YMR_MSGGUID="'+MsgGuid+'" AND YMR_EMAILADDRESS="'+UnDestinataire+'"';
   RSql:=OpenSql (SSql,True);
   if not (RSql.eof) then UnLibelle:=RSql.FindField ('YMR_EMAILNAME').AsString;
   Ferme (RSql);
   if unLibelle<>'' then
    begin
     Result:=UnLibelle;
     Exit;
    end;
  end;

 //--- Recherche des libelles dans la table US_UTILISATEUR
 SSql:='SELECT US_LIBELLE FROM UTILISAT WHERE US_UTILISATEUR="'+UnDestinataire+'"';
 RSql:=OpenSql (SSql,True);
 if not (RSql.eof) then UnLibelle:=RSql.FindField ('US_LIBELLE').AsString;
 Ferme (RSql);
 if unLibelle<>'' then
  begin
   Result:=UnLibelle;
   Exit;
  end;


 //--- Recherche des libelles dans la table YREPERTPERSO
 SSql:='SELECT YRP_NOM, YRP_PRENOM FROM YREPERTPERSO WHERE YRP_EMAIL="'+UnDestinataire+'"';
 RSql:=OpenSql (SSql,True);
 if not (RSql.eof) then UnLibelle:= trim (RSql.FindField ('YRP_NOM').AsString+' '+RSql.FindField ('YRP_PRENOM').AsString);
 Ferme (RSql);
 if unLibelle<>'' then
  begin
   Result:=UnLibelle;
   Exit;
  end;

 //--- Recherche des libelles dans la table YREPERTPERSO
 SSql:='SELECT YLI_LIBELLELISTE FROM YLISTEDISTRIB WHERE YLI_LISTEDISTRIB="'+UnDestinataire+'"';
 RSql:=OpenSql (SSql,True);
 if not (RSql.eof) then UnLibelle:=RSql.FindField ('YLI_LIBELLELISTE').AsString;
 Ferme (RSql);
 Result:=UnLibelle;


end;

//-----------------------------------------
//--- Nom : DonnerListeDestinataire
//-----------------------------------------
function DonnerListeDestinataireFromSql (UnMsgGuid: string; TypeDest : String; ListeDestinataire: string) : String;
var SSql                   : String;
    RSql                   : TQuery;
    UneListeDestinataire   : String;
    UnNom,Unlibelle,UnType : String;
begin
 SSql:='';
 if (TypeDest='A') then SSql := 'SELECT YMR_EMAILNAME,YMR_EMAILADDRESS FROM YMSGADDRESS WHERE YMR_MSGGUID="' + UnMsgGuid + '" AND YMR_ADTYPE=2 ORDER BY YMR_ADORDRE';
 if (TypeDest='Cc') then SSql:= 'SELECT YMR_EMAILNAME,YMR_EMAILADDRESS FROM YMSGADDRESS WHERE YMR_MSGGUID="' + UnMsgGuid + '" AND YMR_ADTYPE=3 ORDER BY YMR_ADORDRE';
 RSql:=OpenSql (SSql,True,-1,'',True);

 if (RSql.eof) then
  Result:=ListeDestinataire
 else
  begin
   UneListeDestinataire:='';
   while not (RSql.eof) do
    begin
     if UneListeDestinataire<>'' then UneListeDestinataire:=UneListeDestinataire+';';
     UnNom:=RSql.FindField ('YMR_EMAILADDRESS').AsString;
     UnLibelle:=RSql.FindField ('YMR_EMAILNAME').AsString;
     UnType:='';
     UneListeDestinataire:=UneListeDestinataire+UnNom+'|'+UnLibelle+'|'+UnType;
     RSql.Next
    end;
   Result:=UneListeDestinataire;
  end;
 Ferme (Rsql);
end;

//-------------------------------------------
//--- Nom : DonnerListeDestinataireFromTob
//-------------------------------------------
function DonnerListeDestinataireFromTob (UneTob : Tob; Detail : Boolean) : String;
var Indice            : Integer;
    ListeDestinataire : String;
begin
 ListeDestinataire:='';

 if (UneTob<>nil) then
  begin
   for Indice:=0 to UneTob.Detail.Count-1 do
    begin
     if (not Detail) then
      begin
       if ListeDestinataire<>'' then ListeDestinataire := ListeDestinataire + ';';
       ListeDestinataire := ListeDestinataire + Unetob.Detail[Indice].GetValue ('Nom');
      end
     else
      begin
       if ListeDestinataire<>'' then ListeDestinataire := ListeDestinataire + ';';
       ListeDestinataire:=ListeDestinataire+Unetob.Detail[Indice].GetValue ('Libelle')+' ('+UneTob.Detail[Indice].GetValue ('Nom')+')';
      end;
    end;
  end;

 Result:=ListeDestinataire;
end;

//-------------------------------------------
//--- Nom : DonnerElementListeDistribution
//-------------------------------------------
function DonnerElementListeDistribution (UneListeDistribution : String) : String;
var SSql                     : String;
    RSql                     : TQuery;
    ElementListeDistribution : String;
begin
 Result:='';ElementListeDistribution:='';
 if not (ExisteSql ('SELECT 1 FROM YLISTEDISTRIB WHERE YLI_LISTEDISTRIB="'+UneListeDistribution+'"')) then
  PGIInfo('Impossible de trouver la liste de distribution '+UneListeDistribution, '')
 else
  begin
   SSql:='SELECT YLA_EMAILADDRESS FROM YLISTEDISTRIBADR WHERE YLA_LISTEDISTRIB="'+UneListeDistribution+'" ORDER BY YLA_EMAILADDRESS';
   RSql:=OpenSql (SSql,True,-1,'',True);
   while not RSql.Eof do
    begin
     ElementListeDistribution:=ElementListeDistribution+RSql.FindField ('YLA_EMAILADDRESS').AsString+';';
     RSql.Next;
    end;
   Result:=ElementListeDistribution;
  end;
 Ferme (RSql);
end;

//-----------------------------------------
//--- Nom : SupprimerDestinataireDansTob
//-----------------------------------------
procedure SupprimerDestinataireDansTob (UneTob : Tob; UneListeDestinataire : String);
var Indice            : Integer;
    UnDestinataire    : String;
    ListeDestinataire : String;
    TrouverDest       : Boolean;
begin
 if (UneTob<>nil) and (UneListeDestinataire<>'') then
  begin
   for Indice:=UneTob.Detail.Count-1 Downto 0 do
    begin
     TrouverDest:=False;
     ListeDestinataire:=UneListeDestinataire;
     while (ListeDestinataire<>'') and (not TrouverDest) do
      begin
       UnDestinataire:=ReadToKenPipe (ListeDestinataire,';');
       if (UneTob.Detail [Indice].Getvalue ('Nom')=UnDestinataire) then TrouverDest:=True;
      end;

     if (not TrouverDest) then UneTob.detail[Indice].Free;
    end;
  end;
end;

//-----------------------------------------
//--- Nom : AjouterDestinataireDansTob
//-----------------------------------------
procedure AjouterDestinataireDansTob (UneTob : Tob; UneListeDestinataire : String);
var UnDestinataire      : String;
    UnLibelle           : String;
    UneInfoDestinataire : String;
    UneTobEnreg         : Tob;
begin
 if (UneListeDestinataire<>'') then
  begin
   while (UneListeDestinataire<>'') do
    begin
     UnDestinataire:=ReadTokenPipe (UneListeDestinataire,';');
     if (UneTob.FindFirst (['Nom'],[UnDestinataire],True)=nil) then
      begin
       UnLibelle:=DonnerLibelleDestinataire ('',UnDestinataire);
       UneInfoDestinataire:=UnDestinataire+'|'+UnLibelle+'|';
       UneTobEnreg:=Tob.Create ('',UneTob,-1);
       UneTobEnreg.LoadFromSt ('Nom|Libelle|Type','|',UneInfoDestinataire,'|');
      end
    end;
  end;
end;

end.

