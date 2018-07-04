unit InitStdcompta;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, hmsgbox, Mask, Hctrls, StdCtrls, Menus, DB, DBTables, Hqry, Grids,
  DBGrids, ExtCtrls, ComCtrls, Buttons, HEnt1, Ent1, General, HStatus, RapSuppr,
  HCompte, HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HPanel,UiUtil, HRichOLE,
  M3FP,MulGene,MZSUtil,QRPlanGe, QRTIER, QRSECTIO, QRPLANJO,Tiers,Fiche,
  Section,Journal,Paramsoc;


implementation

procedure MajGVentil;
var Q1        : Tquery;
begin
  Q1 := Opensql ('SELECT *From GENERAUX', TRUE);
  while not Q1.EOF do
  begin
  if(Q1.FindField('G_VENTILABLE1').asstring = 'X') or
  (Q1.FindField('G_VENTILABLE2').asstring = 'X') or
  (Q1.FindField('G_VENTILABLE3').asstring = 'X') or
  (Q1.FindField('G_VENTILABLE4').asstring = 'X') or
  (Q1.FindField('G_VENTILABLE5').asstring = 'X') then
  begin
  ExecuteSQL('UPDATE  GENERAUX SET G_VENTILABLE="X"'+' where G_GENERAL="'+ Q1.FindField('G_GENERAL').asstring+'"') ;
  if (Q1.FindField('G_GENERAL').asstring = Getparamsoc ('SO_ECCEUROCREDIT'))
    or (Q1.FindField('G_GENERAL').asstring = Getparamsoc ('SO_ECCEURODEBIT')) then
    ExecuteSQL('UPDATE  GENERAUX SET G_VENTILABLE="-",G_VENTILABLE1="-",'+
    'G_VENTILABLE2="-",G_VENTILABLE3="-",G_VENTILABLE4="-",G_VENTILABLE5="-" where G_GENERAL="'+ Q1.FindField('G_GENERAL').asstring+'"') ;
  end;
  Q1.Next;
  end;
  Ferme (Q1);
end;

procedure AglCreationCompte (parms: array of variant; nb: integer);
begin
if parms[0] = 'GENERAUX' then
FicheGene(Nil,'','',taCreatEnSerie,0) ;
if parms[0] = 'TIERS' then
FicheTiers(Nil,'','',taCreatEnSerie,0) ;
if parms[0] = 'SECTION' then
FicheSection(Nil,'A1','',taCreatEnSerie,0) ;
if parms[0] = 'JOURNAL' then
FicheJournal(Nil,'','',taCreatEnSerie,0) ;

end;

procedure AglModifCompte (parms: array of variant; nb: integer);
var
  F : TForm;
  Liste   : THDBGrid;
  Query   : THQuery;
  Sav     : string;
begin
  F:=TForm(Longint(Parms[0])) ;
  if F=Nil then exit ;
  Liste:=THDBGrid(F.FindComponent('FListe') );
  if Liste=Nil then exit;
  Query:=THQuery(F.FindComponent('Q')) ;
  if (Query=Nil) then exit;
  if (parms[2] = 'Modif') then
  begin
       if parms[3] ='GENERAUX' then
          FicheGene(Query,'',parms[1],taModif,0) ;
       if parms[3] ='TIERS' then
          FicheTiers(Query,'',parms[1],taModif,0) ;
       if parms[3] ='JOURNAL' then
          FicheJournal(Query,'',parms[1],taModif,0) ;
       if parms[3] = 'SECTION' then
          FicheSection(Query,parms[4],parms[1],taModif,0) ;
     end;
  if (parms[2] = 'Serie') then
  begin
             if (Liste.NbSelected>0) or (Liste.AllSelected) then
              BEGIN
              // parms[3] ='GENERAUX' , 'TIERS'
              if parms[3] = 'SECTION' then
                 ModifieEnSerie(parms[3],parms[4],Liste,Query)
              else
              begin
                   Sav := V_PGI.ExtendedFieldSelection;
                   if (Sav = 'A') and (GetParamSoc('SO_CPPCLSANSANA')=True) then
                   begin
                         PGIInfo('L''option analytique n''est pas activée','') ;
                         exit;
                   end;
                   ModifieEnSerie(parms[3],'',Liste,Query) ;
                   if Sav = 'A' then  // pour les ventilations ana
                   begin
                        V_PGI.ExtendedFieldSelection := Sav;
                        MajGVentil;
                   end;
              end;
              END ;

  end;
end;

procedure AglEditCompte (parms: array of variant; nb: integer);
begin
     if parms[0] = 'GENERAUX' then
        PlanGeneral('', false);
     if parms[0] = 'TIERS' then
        PlanAuxi('',False) ;
     if parms[0] = 'SECTION' then
        PlanSection('','',FALSE) ;
     if parms[0] = 'JOURNAL' then
        PlanJournal('',FALSE) ;
end;
procedure AGLValiderNum(parms: array of variant; nb: integer );
var
  F        : TForm;
begin
  F:=TForm(Longint(Parms[0])) ;
  TFFiche(F).ModalResult := 1;
end;


Initialization
  RegisterAglProc( 'CreationCompte', FALSE , 1, AglCreationCompte);
  RegisterAglProc( 'ModifCompte', TRUE , 1, AglModifCompte);
  RegisterAglProc( 'EditCompte', FALSE , 1, AglEditCompte);
  RegisterAglProc( 'ValiderNum', TRUE , 1, AGLValiderNum);

finalization

end.
