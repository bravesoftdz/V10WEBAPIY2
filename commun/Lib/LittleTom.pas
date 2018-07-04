(*
LM20060516 :  facilite l'embarquement/gestion de toms annexes à la tom en cours
              en prenant en charge la mise à jour d'enregistrements
              de tables autres que la table de la tom
*)
unit LittleTom;

interface

uses
   SysUtils, Controls,
   db, HTB97, hctrls, StdCtrls, dpOutils, ComCtrls,
   Classes, utob, forms,  EventDecla, hent1, uTom
   ;

type TPDataset = ^TDataset;

type
  TLittleTOM = class(TComponent)
  private
    Table : string ;
    Ecran : TForm ;

    PDS_c : TPDataset ;//de la tom "principale" ...
    LaTom : Tom; //la tom "principale" ...
    peutSauver : boolean ;
    FExclusVerifModif : string ;

    function  Modifie : boolean ;
    procedure AffecteCle(ValCle:string) ;
    procedure initExitZone ;
    procedure initClickZone ;

    procedure SetExclusVerifModif(champ:string) ;
    function casInsertSimple (ValCle:string) : boolean ;

  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;

  published
    EvDecla : TEventDecla ;
    procedure OnClick(sender:TObject) ;
    procedure OnExit (sender:TObject) ;

  public
    tb_ : Tob ;


    constructor create (f:TForm; DbTable:string; UneTom:Tom; PDataset_p : TPDataset); overload; //LM20071105
    destructor  Destroy; override;//LM20071105

    function  ChargeEnreg(valCle:string) :boolean;
    procedure NewEnreg  ;
    function  sauve(valCle:string) :boolean;
    function  get(nomchamp:string):variant ;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : 
Créé le ...... : 12/11/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
    procedure put(nomchamp:string; valeur:variant );

    property  IsModifie : boolean read Modifie ;
    property  ExclusVerifModif : string read FExclusVerifModif  write SetExclusVerifModif ;
 end ;

implementation
uses uSatUtil;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : 
Créé le ...... : 12/11/2007
Modifié le ... : 12/11/2007
Description .. : Pointeur sur DataSet :
Suite ........ : - en 2 tiers le DataSet est chargé dans le OnArgument
Suite ........ : - en EAGL le DataSet est chargé entre le OnArgument et le 
Suite ........ : OnLoad
Mots clefs ... : 
*****************************************************************}
constructor TLittleTOM.create (f:TForm; DbTable:string; uneTom:Tom; PDataset_p : TPDataset);
begin
  inherited Create (f) ;  //LM20071105

  Table:=DbTable ;
  tb_ := Tob.Create(table, nil, -1) ;
  Ecran:=f ;
  if Ecran <> nil then Ecran.FreeNotification(Self);
  EvDecla := TEventDecla.Create(f) ;
  PDS_c := PDataSet_p;
  LaTom:=uneTom ;
  initExitZone;
  initClickZone;
end ;

destructor TLittleTOM.Destroy;
begin
  tb_.Free ;
  //EvDecla.free ; //MD20070705 assuré par le owner f
  inherited ;             //LM20071105 
end ;


procedure TLittleTOM.AffecteCle(ValCle : string) ;
var s, cle : string ;
    i:integer ;
begin
  i:=1;
  s := TableToCle1(Table) ;
  while s<>'' do
  begin
    cle := ReadTokenstV (s) ;
    tb_.P(cle, gtfs(valCle,'|',i));//pipe en séparateur en cas de +sieurs champs dans la clé
    inc(i) ;
  end ;
end ;


function TLittleTOM.ChargeEnreg(ValCle:string) :boolean;
begin
  result:=false;
  tb_.InitValeurs;
  if ValCle <> '' then
  begin
    AffecteCle(ValCle) ;
    result:=tb_.LoadDb;
  end ;
  if (Ecran<>nil) and (result = true) then
      tb_.PutEcran(Ecran) ;
  peutSauver :=result ;

end ;

procedure TLittleTOM.NewEnreg;
begin
  tb_.InitValeurs;
  if (Ecran<>nil)then
      tb_.PutEcran(Ecran) ;
end ;

function TLittleTOM.Sauve(valCle:string) :boolean;
var ok : boolean ;
begin
  result:=false ;
  if (Ecran=nil) then exit ;
  tb_.getEcran(Ecran);
  affecteCle(valCle) ;
  Ok := peutSauver or (not peutSauver and casInsertSimple(valCle));
  if Ok then result:=tb_.InsertOrUpdateDB;
end ;

function  TLittleTOM.get(nomchamp:string):variant ;
begin
  result := tb_.g(nomchamp);
end ;

procedure TLittleTOM.put(nomchamp:string; valeur:variant );
begin
  tb_.p(nomchamp, valeur)
end ;

function TLittleTOM.Modifie : boolean ;
var w_:Tob ;
    colName : string ;
    i : integer ;
begin
  result:=false ;
  if Ecran = nil then exit ;
   w_:=tob.create(table, nil,-1) ;
  w_.Dupliquer(tb_, true, true);
  w_.getEcran(Ecran) ;

  for i:=1 to w_.NbChamps do
    if pos(w_.GetNomChamp(i)+',', FExclusVerifModif) = 0 then
      colName := colName + w_.GetNomChamp(i) + ';' ;

  result := compareTob (w_, tb_, colName) <> 0 ;
  w_.free ;
end ;

procedure TLittleTOM.initExitZone ;
var i : integer ;
    cp : TComponent ;
    nom: string ;
begin
  for i:=0 to Ecran.ComponentCount-1 do
  begin
    nom:= upperCase(Ecran.Components[i].name) ;
    if (champToNum(nom)>-1) and (ExtractPrefixe(nom)=TableToPrefixe(Table)) then
      EvDecla.rebranche (nom, 'OnExit', OnExit ) ;
  end ;
end ;

procedure TLittleTOM.onExit (sender:TObject) ;
var index, nom, st : string ;
begin
  if (sender=nil)
   or (table='') 
    or (csDestroying in Ecran.componentState) then exit ;
  nom:=TControl(sender).name ;
  if champToNum(nom)=0 then exit ;
  EvDecla.Exec(nom, 'OnExit', sender);

  if Modifie and (PDS_c^ <> nil) then
  begin
    if not (PDS_c^.State in [dsInsert, dsEdit]) then PDS_c^.Edit;
    index := TableToCle1(latom.TableName);//Table) ;
    if pos(',', index)>0 then index := ReadTokenstV (index) ;

    if (index<>'') and (LaTom<>nil) then
      laTom.setField(index, laTom.GetField(index)) ;
  end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 12/11/2007
Modifié le ... :   /  /
Description .. : Pour gestion des checkbox
Mots clefs ... :
*****************************************************************}
procedure TLittleTOM.initClickZone;
var i : integer ;
    cp : TComponent ;
    nom: string ;
begin
   for i:=0 to Ecran.ComponentCount-1 do
   begin
//      if not (Ecran.Components[i] is TCheckBox) then continue;
      nom:= upperCase(Ecran.Components[i].name) ;
      if (champToNum(nom)>-1) and (ExtractPrefixe(nom)=TableToPrefixe(Table)) then
         EvDecla.rebranche (nom, 'OnClick', OnClick) ;
  end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 12/11/2007
Modifié le ... :   /  /
Description .. : Pour gestion des checkbox
Mots clefs ... :
*****************************************************************}
procedure TLittleTOM.OnClick(sender:TObject) ;
var index, nom, st : string ;
begin
  if (sender=nil) or (table='') or (csDestroying in Ecran.componentState) then exit ;
//  if not (TControl(sender) is TCheckBox) then exit;

  nom:=TControl(sender).name ;
  if champToNum(nom)=0 then exit ;
  EvDecla.Exec(nom, 'OnClick', sender);

  if Modifie and (PDS_c^<>nil) then
  begin
    if not (PDS_c^.State in [dsInsert, dsEdit]) then
      PDS_c^.Edit;
    index := TableToCle1(latom.TableName);//Table) ;

    if pos(',', index)>0 then
      index := ReadTokenstV (index) ;

    if (index<>'') and (LaTom<>nil) then
      laTom.setField(index, laTom.GetField(index)) ;
  end ;
end ;

procedure TLittleTOM.Notification(AComponent: TComponent; Operation: TOperation);
begin
  if (Operation = opRemove) and (AComponent = Ecran) then
      Ecran := nil ;
  inherited;
end;

procedure TLittleTOM.SetExclusVerifModif(champ:string) ;
begin
  FExclusVerifModif:=stringReplace(champ, ' ' , '', [rfReplaceAll]) ;
  FExclusVerifModif:=stringReplace(FExclusVerifModif, ';' , ',', [rfReplaceAll]) ;
  if FExclusVerifModif[length(FExclusVerifModif)]<>',' then
    FExclusVerifModif:=FExclusVerifModif+',' ;
end ;


function TLittleTOM.casInsertSimple (ValCle:string) : boolean ;
var champ, s, cle : string ;
    i : integer ;
begin
  result:=false ;
  i:=0 ;
  s := TableToCle1(Table) ;
  while s<>'' do
  begin
    cle := ReadTokenstV (s) ;
    if cle<>'' then
    begin
      if i=0 then champ:=cle ;
      inc(i) ;
    end
    else
      break ;
  end ;

  if i>1 then exit ; //clé composée non gérée. Pour le moment index = guid

  if champ<>'' then
    result:=not ExisteSQL (  'select 1 from ' + Table +
                            ' where ' + champ +' = "' + tb_.g(champ) + '"' ) ;

end ;



end.
