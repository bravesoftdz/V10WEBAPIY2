{***********UNITE*************************************************
Auteur  ...... :  NA
Créé le ...... : 06/06/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGANALYSESIMUL ()
                 Analyse des simulations
Mots clefs ... : TOF;PGANALYSESIMUL
*****************************************************************}
Unit PGANALYSESIMUL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     sysutils,
     ComCtrls,
     P5def,
     pgoutils2,
     Hctrls,
     UTOF ;

Type
  TOF_PGANALYSESIMUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    procedure verifsimul(sender:TObject);
  end ;

Implementation

procedure TOF_PGANALYSESIMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSESIMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSESIMUL.OnUpdate ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/06/2006
Modifié le ... :   /  /    
Description .. : ON ARGUMENT
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGANALYSESIMUL.OnArgument (S : String ) ;
VAR
num : integer;
nodossier, pred, sql1,st : string;
codesimul : THValComboBox;

begin
  Inherited ;
 // Recherche des libellés des zones TRAVAIL, CODESTAT, LIBREPCMB
  for Num := 1 to 4 do
  begin
  VisibiliteChampSalarie(IntToStr(Num), GetControl('PSR_TRAVAILN' + IntToStr(Num)), GetControl('TPSR_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PSR_CODESTAT'), GetControl('TPSR_CODESTAT'));

  for Num := 1 to 4 do
  begin
  VisibiliteChamplibresal(Inttostr(Num), GetControl('PSR_LIBREPCMB'+ Inttostr(Num)),GetControl('TPSR_LIBREPCMB'+ IntToStr(Num)));
  end;

  // Initialise le plus piur la tablette PGPOPULATIONIDR
   nodossier := PgrendNodossier();
  // recherche si population existe dans le dossier
  pred := 'DOS';
  sql1 := 'select PPC_POPULATION,PPO_CODEPOP from ORDREPOPULATION,CODEPOPULATION WHERE PPC_CODEPOP = PPO_CODEPOP'+
  ' and PPO_PREDEFINI = "DOS" and PPO_NODOSSIER = "'+nodossier+'" and PPO_TYPEPOP like "%IDR%"';
  if not existesql(sql1) then
  begin
     // recherche si population existe en standard
     pred := 'STD';
    sql1 := 'select PPC_POPULATION,PPO_CODEPOP from ORDREPOPULATION, CODEPOPULATION WHERE PPC_CODEPOP = PPO_CODEPOP'+
    ' and PPO_PREDEFINI = "STD" and and PPO_TYPEPOP like "%IDR%"';
    if not existesql(sql1) then   pred := 'CEG';      // sinon pop CEGID

  end;
  st := ' and PPC_PREDEFINI = "'+pred+'"';
  setcontrolproperty('PSR_POPULATION', 'plus', st);

  // Recherche di le code simulation est renseigné pour afficher les pop de cette simulation
  Codesimul := THValcomboBox(Getcontrol('PSR_CODESIMUL'));
  if Assigned(codesimul) then codesimul.Onchange := verifsimul;

end ;

procedure TOF_PGANALYSESIMUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSESIMUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSESIMUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSESIMUL.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : Recherche des populations de la simulation 
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGANALYSESIMUL.verifsimul(sender : Tobject);
var
codesimul , sql1, popul, listewhere,st,selectpop : string;
qq: Tquery;

begin

codesimul := THValcomboBox(getcontrol('PSR_CODESIMUL')).value;
  if codesimul <> '' then
  begin
    // Recherche des populations traités dans la simulation
    sql1 := 'select PSL_SELECTPOP FROM SIMULATION WHERE PSL_CODESIMUL = "'+codesimul+'"';
    QQ:= opensql(sql1,true);
    if not QQ.EOF then
    begin
      selectpop := QQ.findfield('PSL_SELECTPOP').asstring;
      popul := READTOKENST(selectpop);
      while popul <> '' do                                // lecture de chaque population
      begin
      listewhere :=  listewhere + '"'+popul+'" ,' ;
      popul := READTOKENST(selectpop);
      end;

      if listewhere <> '' then
      begin                // si au moins une population
      listewhere := '(' + copy(listewhere, 1, length(listewhere) - 1) + ')';

      St := 'AND PPC_POPULATION in  '+listewhere+'';
      setcontrolproperty('PSR_POPULATION', 'plus', st);
      end;
    end;
  end;

end;
Initialization
  registerclasses ( [ TOF_PGANALYSESIMUL ] ) ; 
end.
