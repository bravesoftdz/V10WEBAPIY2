unit GcArticle_Tof;

interface

uses  Classes,HDimension,sysutils,HEnt1,HCtrls,Menus,Math,
{$IFDEF EAGLCLIENT}
      eFiche,Maineagl,
{$ELSE}
      FE_Main,DBCtrls,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fiche,
{$ENDIF}
      Forms,Utom,UTof,AGLInit;

Type
    TOF_GcArticle = Class (TOF)
      private
        GA_PRIXPOURQTE : THEdit ;
        PPQ_Old : Double;
        function  NbreDecimalSupp: Integer;
        procedure FormatZone ;
        procedure GA_PRIXPOURQTEOnExit(Sender : TObject) ;
        procedure AppelGestionQte(Sender: TObject);
      public
        procedure OnArgument(StArg : string);  override;
        procedure OnLoad;  override;
    end;

implementation

Uses UtomArticle ;


procedure TOF_GcArticle.OnArgument(StArg : string);
Var OM : TOM;
     PComplement : TPopupMenu ;
     MnQte : TMenuItem;
begin
if (copy(Ecran.Name,1,9)<>'GCARTICLE') then exit;
if (Ecran is TFFiche) then OM:=TFFiche(Ecran).OM else exit;
if Not(OM is TOM_Article) then exit;
inherited;
GA_PRIXPOURQTE:=THEdit(OM.GetControl('GA_PRIXPOURQTE'));
GA_PRIXPOURQTE.OnExit:=GA_PRIXPOURQTEOnExit;
PComplement:=TPopupMenu(OM.GetControl('POPCOMPLEMENT'));
MnQte:=TMenuItem.Create(PComplement);
MnQte.Caption:=TraduireMemoire('Gestion quantité');
MnQte.Name:='MnQte';
MnQte.OnMeasureItem:=PComplement.Items[0].OnMeasureItem;
MnQte.OnDrawItem:=PComplement.Items[0].OnDrawItem;
MnQte.OnClick:=AppelGestionQte;
PComplement.Items.Add(MnQte);
end;

procedure TOF_GcArticle.OnLoad;
begin
if (copy(Ecran.Name,1,9)<>'GCARTICLE') then exit;
inherited;
FormatZone;
end;

function TOF_GcArticle.NbreDecimalSupp: Integer;
Var PPQ : Double;
    NbDec : integer;
begin
NbDec:=0;
PPQ:=Valeur(GetControlText('GA_PRIXPOURQTE')); if PPQ<=0 then PPQ:=1 ;
While PPQ>1 do
  begin
  PPQ:=arrondi((PPQ/10.0)-0.499,0);
  inc(NbDec);
  end;
Result:=NbDec;
end;

procedure TOF_GcArticle.FormatZone ;
Var OM : TOM;
    FF : string ;
    ii : integer ;
begin
if (Ecran is TFFiche) then OM:=TFFiche(Ecran).OM else exit;
if Not(OM is TOM_Article) then exit;
PPQ_Old:=Valeur(GetControlText('GA_PRIXPOURQTE'));
FF:='#,##0.';
for ii:=1 to Max(V_PGI.OkDecP, V_PGI.OkDecV+NbreDecimalSupp) do FF:=FF+'0';
OM.SetControlProperty('GA_PAHT','DisplayFormat',FF);
OM.SetControlProperty('GA_DPA','DisplayFormat',FF);
OM.SetControlProperty('GA_DPR','DisplayFormat',FF);
OM.SetControlProperty('GA_PMAP','DisplayFormat',FF);
OM.SetControlProperty('GA_PMRP','DisplayFormat',FF);
end;

procedure TOF_GcArticle.GA_PRIXPOURQTEOnExit(Sender : TObject) ;
begin
if PPQ_Old<>Valeur(GetControlText('GA_PRIXPOURQTE'))then FormatZone;
end;

procedure TOF_GcArticle.AppelGestionQte(Sender: TObject);
Var ActionFich,Range,Lequel : string;
    OM : TOM;
begin
if (Ecran is TFFiche) then OM:=TFFiche(Ecran).OM else exit;
if Not(OM is TOM_Article) then exit;
if TFFiche(Ecran).TypeAction <> taConsult then
begin
    if ExisteSql('Select GAF_ARTICLE From ArticleQte Where GAF_ARTICLE="'+OM.GetField('GA_ARTICLE')+'"') then
    begin
      ActionFich:=ActionToString(taModif);
      Range:='';
      Lequel:=OM.GetField('GA_ARTICLE');
      end else
      begin
      ActionFich:=ActionToString(taCreat);
      Lequel:='';
      Range:=OM.GetField('GA_ARTICLE');
    end;
end else ActionFich:=ActionToString(taConsult);
AGLLanceFiche('GC','GCARTICLEQTE',Range,Lequel,ActionFich);
end;

Initialization
registerclasses([TOF_GcArticle]) ;

end.
