unit AGlStdcompta;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, hmsgbox, Mask, Hctrls, StdCtrls, Menus, DB, DBTables, Hqry, Grids,
  DBGrids, ExtCtrls, ComCtrls, Buttons, HEnt1, Ent1, General, HStatus, RapSuppr,
  HCompte, HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HPanel,UiUtil, HRichOLE,
  M3FP,MulGene,QRPlanGe, QRTIER, QRSECTIO, QRPLANJO,Tiers,Fiche,
{$IFDEF CCS3}
{$ELSE}
  MSUtil,
{$ENDIF}
  Section,Journal,Paramsoc;


implementation

uses
     SaisUtil, Letbatch,
     Spin, Region, ed_tools;

// ajout me 30/04/01
Function AnaEstVentilable(Quelaxe : Byte; Cpte : string) : Boolean ;
var
Q1       : TQuery;
BEGIN
Result:=False ;
Q1 := Opensql ('SELECT G_VENTILABLE1,G_VENTILABLE2,G_VENTILABLE3,G_VENTILABLE4,G_VENTILABLE5 From GENERAUX Where G_GENERAL="'+Cpte+'"', TRUE);
if not Q1.EOF then
begin
     Case QuelAxe of
       0: if (Q1.FindField('G_VENTILABLE1').asstring='X') or
              (Q1.FindField('G_VENTILABLE2').asstring='X') or
              (Q1.FindField('G_VENTILABLE3').asstring='X') or
              (Q1.FindField('G_VENTILABLE4').asstring='X') or
              (Q1.FindField('G_VENTILABLE5').asstring='X') then Result:=True;
       1: if Q1.FindField('G_VENTILABLE1').asstring='X' then Result:=True ;
       2: if Q1.FindField('G_VENTILABLE2').asstring='X' then Result:=True ;
       3: if Q1.FindField('G_VENTILABLE3').asstring='X' then Result:=True ;
       4: if Q1.FindField('G_VENTILABLE4').asstring='X' then Result:=True ;
       5: if Q1.FindField('G_VENTILABLE5').asstring='X' then Result:=True ;
    End ;
end;
ferme (Q1);
END ;

procedure MajGVentil(OKgene : Boolean; LeCpte : string);
var Q1,Q2,Qana        : Tquery;
    Dev               : RDevise ;
    OEcr,OAna         : TOBM ;
    TVentAna          : TList ;
    NewAxe            : String ;
    i ,k              : Integer ;
    Okok              : Boolean ;
    Where            : string;
begin

if OkGene then
begin
 Where := 'Where (G_VENTILABLE1="X" or G_VENTILABLE2="X"'+
 'or G_VENTILABLE3="X" or G_VENTILABLE4="X" or G_VENTILABLE5="X")';
 if ExisteSQL ('SELECT * From GENERAUX '+ Where) then
    ExecuteSQL('UPDATE  GENERAUX SET G_VENTILABLE="X" '+Where) ;
 Where := 'Where  (G_GENERAL="'+ Getparamsoc ('SO_ECCEUROCREDIT')+'" or '+
      'G_GENERAL="'+ Getparamsoc ('SO_ECCEURODEBIT')+ '")' ;
 if ExisteSQL ('SELECT * From GENERAUX '+ Where) then
      ExecuteSQL('UPDATE  GENERAUX SET G_VENTILABLE="-",G_VENTILABLE1="-",'+
      'G_VENTILABLE2="-",G_VENTILABLE3="-",G_VENTILABLE4="-",G_VENTILABLE5="-" ' + Where) ;
end
else
begin
  Okok:=False ;
  Q1:=OpenSQL('Select * from ECRITURE Where E_ANA<>"X" AND E_GENERAL="'+LeCpte+'" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")',True) ;
  QAna:=OpenSQL('Select * from ANALYTIQ WHERE Y_SECTION="RerER"',False) ;
  OEcr:=TOBM.Create(EcrGen,'',False) ; TVentAna:=TList.Create ; Okok:=False ;
  While not Q1.EOF do
   BEGIN
   OEcr.ChargeMvt(Q1) ;
   Dev.Code:=Q1.FindField('E_DEVISE').AsString ; GetInfosDevise(Dev) ;
   Dev.DateTaux:=Q1.FindField('E_DATETAUXDEV').AsDateTime ;
   Dev.Taux:=Q1.FindField('E_TAUXDEV').AsFloat ;
   Okok:=False ;
   for i:=1 to MaxAxe do if AnaEstVentilable(i,LeCpte) then
       BEGIN
       VideListe(TVentAna) ; NewAxe:='A'+IntToStr(i) ; Okok:=True ;
       VentileGenerale(NewAxe,OEcr,Dev,TVentAna,True) ;
       for k:=0 to TVentAna.Count-1 do
           BEGIN
           QAna.Insert ;
           OAna:=TOBM(TVentAna[k]) ; OAna.EgalChamps(QAna) ;
           QAna.Post ;
           MajSoldeSection(QAna,True) ;
           END ;
       END ;
    Q1.Next ;
   END ;
   if Okok then
      ExecuteSQL('UPDATE ECRITURE SET E_ANA="X" WHERE E_GENERAL="'+LeCpte+'" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")')
   else
   begin // pas de ventilation
      if not AnaEstVentilable(0,LeCpte) then
      begin
       ExecuteSQL('UPDATE GENERAUX SET G_VENTILABLE="-"'+' where G_GENERAL="'+ LeCpte+'"') ;
       ExecuteSQL('UPDATE ECRITURE SET E_ANA="-" WHERE E_GENERAL="'+LeCpte+'" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")') ;
       ExecuteSQL ('DELETE FROM ANALYTIQ Where Y_GENERAL="'+LeCpte+'"');
       ExecuteSQL ('DELETE FROM VENTIL Where V_COMPTE="'+LeCpte+'"');
       end;
   end;
   Ferme(Q1) ; Ferme(QAna) ; VideListe(TVentAna) ; TVentAna.Free ; OEcr.Free ;
end;
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
  i       : integer;
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
{$IFDEF CCS3}
{$ELSE}
                 ModifieEnSerie(parms[3],parms[4],Liste,Query)
{$ENDIF}
                 else
                 begin
                   Sav := V_PGI.ExtendedFieldSelection;
                   if (Sav = '1') and (GetParamSoc('SO_CPPCLSANSANA')=True) then
                   begin
                         PGIInfo('L''option analytique n''est pas activée','') ;
                         exit;
                   end;
                   // voir pour les autres axes 3 à 5 si nécessaire
                   for i:=1 to 2 do if VH^.Cpta[AxeToFb('A'+IntToStr(i))].Attente=''
                   then
                   begin
                         PGIInfo('Attention Axe '+IntToStr(i)+ ' ne possède aucun code section et n''a pas de section d''attente' ,'') ;
                         if i=2 then exit;
                   end;
{$IFDEF CCS3}
{$ELSE}
                   ModifieEnSerie(parms[3],'',Liste,Query) ;
{$ENDIF}
                   if Sav = '1' then  // pour les ventilations ana
                   begin
                        V_PGI.ExtendedFieldSelection := Sav;
                        MajGVentil (TRUE, '');
                        if Liste.AllSelected then
                        BEGIN
                             MoveCur(False);
                             Query.First;
                             while Not Query.EOF do
                             BEGIN
                                  MajGVentil (FALSE, Query.FindField ('G_GENERAL').asstring);
                                  Query.next;
                             END;
                             Liste.AllSelected:=False;
                        END
                        ELSE
                        BEGIN
                             InitMove(Liste.NbSelected,'');
                             for i:=0 to Liste.NbSelected-1 do
                             BEGIN
                             MoveCur(False);
                             Liste.GotoLeBookmark(i);
                             MajGVentil (FALSE, Query.FindField ('G_GENERAL').asstring);
                             END;
                             Liste.ClearSelected;
                             FiniMove ;
                        END;
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
