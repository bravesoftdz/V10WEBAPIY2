{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 09/09/2002
Modifié le ... :   /  /
Description .. : Unité de gestion du multicritère MODIFCOEFF
Mots clefs ... : PAIE,PGMODIFCOEFF
*****************************************************************}
{
PT1   : 20/02/2003 VG V_42 Cas du double-click sur une liste vide
}
unit UTofPG_MulModifCoeff;

interface
uses UTOF,Hqry,AGLInit,HCtrls,classes,
{$IFNDEF EAGLCLIENT}
     HDB,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul,FE_Main,
{$ELSE}
     EMul,MainEAgl,UtileAgl,
{$ENDIF}
     sysutils,HTB97,UTOB,Hent1,hmsgbox;


Type
     TOF_PGMULMODIFCOEFF= Class (TOF)
       public
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnUpdate ; override;

       private
       Q_Mul:THQuery ;  // Query pour changer la liste associee
{$IFNDEF EAGLCLIENT}
       Liste : THDBGrid;
{$ELSE}
       Liste : THGrid;
{$ENDIF}
       BCherche: TToolbarButton97;

       procedure GrilleDblClick (Sender: TObject);
       procedure NewCoeff(Sender: TObject);
     END ;

implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 09/09/2002
Modifié le ... :   /  /
Description .. : OnArgument
Mots clefs ... : PAIE;PGMODIFCOEFF
*****************************************************************}
procedure TOF_PGMULMODIFCOEFF.OnArgument(Arguments: String);
var
St, StDelete : string;
Q : TQuery;
MajCoeff, MaxFin, NewMajCoeff : TDateTime;
TModifCoeff, TModifCoeffFille : TOB;
begin
inherited ;

MaxFin := IDate1900;
St := 'SELECT MAX(PPU_DATEFIN) AS MAXFIN FROM PAIEENCOURS';
Q := OpenSql(St, True);
if NOT Q.EOF then
   try
   MaxFin := Q.FindField('MAXFIN').AsDateTime;
   except
         on E: EConvertError do
            MaxFin:= IDate1900;
   end;
Ferme(Q);

St := 'SELECT PMC_DATEANCIEN, PMC_DATEANCIENFIN'+
      ' FROM MODIFCOEFF WHERE'+
      ' PMC_MODIFCOEFF = 0';

Q := OpenSql(St, True);
if NOT Q.EOF then MajCoeff := Q.FindField('PMC_DATEANCIEN').AsDateTime
 else  MajCoeff := IDate1900; // PORTAGECWAS
Ferme(Q);

St := 'SELECT PMC_DATEANCIEN'+
      ' FROM MODIFCOEFF WHERE'+
      ' PMC_DATEANCIEN="'+UsDateTime(IDate1900)+'"';
if ((MaxFin <> MajCoeff) or (ExisteSQL(St))) then
   begin
   St := 'SELECT * FROM MODIFCOEFF';
   Q := OpenSql(St, True);
   TModifCoeff := TOB.Create('Mère ModifCoeff', NIL, -1);
   TModifCoeff.LoadDetailDB('MODIFCOEFF','','',Q,False);
   Ferme(Q);
   try
      begintrans;
      StDelete := 'DELETE FROM MODIFCOEFF';
      ExecuteSQL(StDelete) ;
      TModifCoeffFille := TModifCoeff.FindFirst([''],[''],FALSE);
      if (TModifCoeffFille <> nil) then
         begin
         While (TModifCoeffFille <> nil) do
               begin
               NewMajCoeff := PlusMois(MaxFin, -(TModifCoeffFille.GetValue('PMC_NBANCIENSUP')))+1;
               TModifCoeffFille.PutValue('PMC_DATEANCIEN', NewMajCoeff);
               NewMajCoeff := PlusMois(MaxFin, -(TModifCoeffFille.GetValue('PMC_NBANCIEN')));
               TModifCoeffFille.PutValue('PMC_DATEANCIENFIN', NewMajCoeff);
               TModifCoeffFille := TModifCoeff.FindNext([''],[''],FALSE);
               end;
         end
      else
         begin
         TModifCoeffFille := TOB.Create('MODIFCOEFF', TModifCoeff,-1);
         TModifCoeffFille.AddChampSup('PMC_MODIFCOEFF', FALSE);
         TModifCoeffFille.AddChampSup('PMC_NBANCIEN', FALSE);
         TModifCoeffFille.AddChampSup('PMC_NBANCIENSUP', FALSE);
         TModifCoeffFille.AddChampSup('PMC_DATEANCIEN', FALSE);
         TModifCoeffFille.AddChampSup('PMC_DATEANCIENFIN', FALSE);
         TModifCoeffFille.PutValue('PMC_MODIFCOEFF', 0);
         TModifCoeffFille.PutValue('PMC_NBANCIEN', 0);
         TModifCoeffFille.PutValue('PMC_NBANCIENSUP', 0);
         NewMajCoeff := PlusMois(MaxFin, 0);
         TModifCoeffFille.PutValue('PMC_DATEANCIEN', NewMajCoeff);
         TModifCoeffFille.PutValue('PMC_DATEANCIENFIN', NewMajCoeff);
         end;
      TModifCoeff.SetAllModifie (TRUE);
      TModifCoeff.InsertDB(nil,FALSE);
      FreeAndNil(TModifCoeff);

      CommitTrans;
   Except
      Rollback;
      PGIBox ('Une erreur est survenue lors de la mise à jour de la base', 'Coefficient');
      end;
   end;

TFMul(Ecran).BOuvrir.OnClick := GrilleDblClick;
TFMul(Ecran).BInsert.OnClick := NewCoeff;

{$IFNDEF EAGLCLIENT}
Liste := THDBGrid(GetControl('FListe'));
{$ELSE}
Liste := THGrid(GetControl('FListe'));
{$ENDIF}
if Liste <> NIL then
   begin
{$IFNDEF EAGLCLIENT}
   Liste.MultiSelection := False;
{$ENDIF}
   Liste.OnDblClick := GrilleDblClick;
   end;
BCherche:=TToolbarButton97(GetControl('BCherche'));
Q_Mul:=THQuery(Ecran.FindComponent('Q'));
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 09/09/2002
Modifié le ... :   /  /
Description .. : OnUpdate
Mots clefs ... : PAIE;PGMODIFCOEFF
*****************************************************************}
procedure TOF_PGMULMODIFCOEFF.OnUpdate;
begin
inherited ;
TFMul(Ecran).BOuvrir.Enabled := True;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 09/09/2002
Modifié le ... :   /  /
Description .. : Double-click sur la grille
Mots clefs ... : PAIE;PGMODIFCOEFF
*****************************************************************}
procedure TOF_PGMULMODIFCOEFF.GrilleDblClick(Sender: TObject);
var
Coeff : String;
begin
if (Q_Mul <> nil) then
   begin
//PT1
   if (Q_Mul.RecordCount = 0) then
      exit;
//FIN PT1
   Coeff := IntToStr(Q_Mul.FindField ('PMC_MODIFCOEFF').AsInteger);     // DB2
   end;

AGLLanceFiche ('PAY','MODIFCOEFF',  '',Coeff, 'ACTION=MODIFICATION;COE');
if BCherche<>nil then
   BCherche.click;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : Création d'un Coefficient
Mots clefs ... : PAIE;PGMODIFCOEFF
*****************************************************************}
procedure TOF_PGMULMODIFCOEFF.NewCoeff(Sender: TObject);
var
NumOrdre : integer;
QCoeff : TQuery;
begin
NumOrdre:=1;
QCoeff:=OpenSQL('SELECT MAX(PMC_MODIFCOEFF) AS NUMMAX FROM MODIFCOEFF',True);
if (QCoeff <> nil) then
   try
   NumOrdre:=(QCoeff.FindField('NUMMAX').AsInteger)+1;
   except
         on E: EConvertError do
            NumOrdre:= 1;
   end;
Ferme(QCoeff);
AGLLanceFiche ('PAY','MODIFCOEFF',  '',IntToStr(NumOrdre), 'ACTION=CREATION;COE');
if BCherche<>nil then
   BCherche.click;
end;


Initialization
registerclasses([TOF_PGMULMODIFCOEFF]);
end.
