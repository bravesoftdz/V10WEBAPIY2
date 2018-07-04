unit TofExpVersS3;

interface
uses Classes,SysUtils,Stdctrls,Hctrls,HMsgBox,Ent1,Hent1,Utof,Utob,FE_Main,Vierge,HStatus,SaisUtil, uDbxDataSet;
type
  TOF_ExpVersS3 = class(TOF)
  private
    DevisePivot, DeviseAutre : String ;
    procedure ExporteVersS3(Exo,Where,NomFichier : string) ;
    procedure FilleParFille(TheGirl : TOB;NomFichier : string) ;
    procedure AjoutChamp(TheGirl : TOB) ;
    procedure RecupAnalytiq(TheGirl : TOB;TauxDev : double;Devise : string ) ;
    procedure RecupJal(NomFichier : string) ;
    procedure RecupCompte(NomFichier : string) ;
    procedure RecupTiers(NomFichier : string) ;
    procedure RecupSec(NomFichier : string);
    procedure RecupDevise(NomFichier : string);
    procedure PieceParPiece(Where,NomFichier : string) ;
    function  RecupArguments(var TCritere : TStringList;var NomFichier : string) : boolean;
    procedure OnClickElipsis(Sender : TObject) ;
  public
    procedure Onload ; override ;
    procedure OnUpDate; override ;
  end;

  TOF_ExpVersS3_2 = class(TOF)
  public
    procedure Onload ; override ;
    procedure OnUpdate ; override ;    
  end;

const TMsg: array[1..6] of string 	= (
      {01}  '1;Export Vers S3 ;Erreur dans la fenetre;W;O;O;O'
      {02} ,'1;Export Vers S3 : Date Comptable;La borne minimum doit être inferieur à la borne maximum;W;O;O;O'
      {03} ,'1;Export Vers S3 : Journal;La borne minimum doit être inferieur à la borne maximum;W;O;O;O'
      {04} ,'1;Export Vers S3 : Numero de Piece;La borne minimum doit être inferieur à la borne maximum;W;O;O;O'
      {05} ,'1;Export Vers S3 : Code Devise;Les codes devises doivent être differents;W;O;O;O'
      {06} ,'1;Export Vers S3 : Code Devise;Les codes devises doivent être renseignés;W;O;O;O'
            ) ;
      SP_NOMJAL       = 1 ;
      SP_VALJAL       = SP_NOMJAL+1 ;
      SP_COCHE        = SP_VALJAL+1 ;
implementation
{***********A.G.L.***********************************************
Auteur  ...... : PELUD Yann
Créé le ...... : 17/01/2000
Modifié le ... : 17/01/2000
Description .. : Transforme une date JJ/MM/AAAA en AAAAMMJJ
Mots clefs ... : DATE;
*****************************************************************}
function ArrangeDate(DateSrc : string) : String ;
var Year,Month,Day : word ;
begin
if DateSrc<>'' then
  begin
  DecodeDate(StrToDate(DateSrc),Year,Month,Day) ;
  result:=Format('%.4d%.2d%.2d',[Year, Month, Day]) ;
  end ;
end ;
{***********A.G.L.***********************************************
Auteur  ...... : PELUD Yann
Créé le ...... : 17/01/2000
Modifié le ... :   /  /    
Description .. : Ajoute X au NomFichier.txt : NomfichierX.txt
Mots clefs ... : CHAINE
*****************************************************************}
Function ConcatNom(NomFichier : string;nb : integer) : string ;
var tmp : string ; Lng,i : integer ;
begin
Lng:=length(NomFichier) ;
for i:=Lng downto 1 do if NomFichier[i]='.' then break ;
if i=0 then   Result:=NomFichier+IntToStr(nb)
else
  begin
  tmp:=copy(NomFichier,i,Lng-i+1) ;
  system.Delete(NomFichier,i,Lng-i+1) ;
  Result:=NomFichier+'_'+IntToStr(nb)+tmp ;
  end ;
end ;
{+===================================+}
{+                                   +}
{+          TOF : EXPVERSS3          +}
{+                                   +}
{+===================================+}
procedure TOF_ExpVersS3.OnLoad;
var FExo,FEtab,FNatJnl1,FNatJnl2,CAxes : THValComboBox;SaufJnl,Filename,MPivot,MAutre : THEdit ;QQ:TQuery ;
begin
SaufJnl:=THEdit(GetControl('SAUFJAL')) ;
SaufJnl.OnElipsisClick:=OnClickElipsis ;
Fexo:=THValComboBox(GetControl('CEXO')) ; FExo.ItemIndex:=0 ;
FEtab:=THValComboBox(GetControl('FETAB')) ; FEtab.ItemIndex:=0 ;
FNatJnl1:=THValComboBox(GetControl('FNATJNL1')) ; FNatJnl1.ItemIndex:=0 ;
FNatJnl2:=THValComboBox(GetControl('FNATJNL2')) ; FNatJnl2.ItemIndex:=FNatJnl2.Items.Count-1 ;
Caxes:=THValComboBox(GetControl('CAXES')) ; CAxes.ItemIndex:=0 ;
FileName:=THEdit(GetControl('FILENAME')) ; FileName.Text:='ExpS3.Txt' ;
MPivot:=THEdit(GetControl('MPIVOT')) ; MPivot.Text:=V_PGI.DevisePivot ;
QQ:=OpenSQL('SELECT D_DEVISE FROM DEVISE WHERE D_FONGIBLE="X"',True) ;
MAutre:=THEdit(GetControl('MAUTRE')) ; MAutre.Text:=QQ.FindField('D_DEVISE').AsString ;
ferme(QQ) ;
end;

procedure TOF_ExpVersS3.OnClickElipsis(Sender : TObject) ;
var SaufJnl : THEdit ;
begin
SaufJnl:=THEdit(GetControl('SAUFJAL')) ;
SaufJnl.Text:=AGLLanceFiche('CP','EXPS3JAL','','','') ;
end ;

procedure TOF_ExpVersS3.OnUpDate;
var Where,NomFichier : string ; CExo : THValComboBox ;
TCritere,TExercic : TStringList; i,j : integer ; QQ : TQuery ;
begin
TCritere:=TStringList.Create; TExercic:=TStringList.Create;
CExo:=THValComboBox(GetControl('CEXO')) ;
if CExo<>nil then
  begin
  if Cexo.ItemIndex>0 then TExercic.Add(Cexo.Value)
  else
    begin
    QQ:=Opensql('SELECT EX_EXERCICE FROM EXERCICE',True) ;
    While not QQ.Eof do begin TExercic.Add(QQ.FindField('EX_EXERCICE').AsString); QQ.Next ; end ;
    ferme(QQ) ;
    end ;
  end ;
if RecupArguments(TCritere,Nomfichier) then
  begin
  for j:=0 to TExercic.Count-1 do
    begin
    Where:=' WHERE E_EXERCICE="'+TExercic.Strings[j]+'"' ;
    i:=TCritere.Count-1;
    while i>=0 do
      begin
      Where:=Where+' AND '+TCritere.strings[i] ;
      Dec(i) ;
      end;
      ExporteVersS3(TExercic.Strings[j],Where,ConcatNom(NomFichier,j)) ;
    end ;
  end ;
end;

procedure TOF_ExpVersS3.ExporteVersS3(Exo,Where,NomFichier : string) ;
var QQ : TQuery ;FDest: TextFile ; i : integer ;Where2 : string ; Entet : boolean ;
begin
Entet:=True ;
assignfile(FDest, NomFichier) ;
rewrite(FDest) ;
QQ:=Opensql('SELECT SO_SIRET FROM SOCIETE',True) ;
writeln(FDest,QQ.FindField('SO_SIRET').AsString);
ferme(QQ) ;
writeln(FDest,DevisePivot);
writeln(FDest,'PGI1');
QQ:=Opensql('SELECT EX_DATEDEBUT,EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+Exo+'"',True) ;
writeln(FDest,ArrangeDate(QQ.FindField('EX_DATEDEBUT').AsString)+'|'+ArrangeDate(QQ.FindField('EX_DATEFIN').AsString));
ferme(QQ) ;
for i:=1 to 20 do Writeln(FDest,'');
CloseFile(FDest);
QQ:=Opensql('SELECT E_NUMEROPIECE,E_JOURNAL FROM ECRITURE '+Where+'GROUP BY E_NUMEROPIECE,E_JOURNAL ',True) ;
InitMove(QQ.Fields.count,'') ;
While not QQ.Eof do
  begin
  if Entet then
    begin
    RecupCompte(Nomfichier)  ;
    RecupTiers(Nomfichier)  ;
    RecupJal(Nomfichier)  ;
    RecupSec(Nomfichier);
    RecupDevise(Nomfichier);
    end ;
  Where2:='WHERE E_NUMEROPIECE="'+QQ.FindField('E_NUMEROPIECE').AsString+'"' ;
  Where2:=Where2+' AND E_JOURNAL="'+QQ.FindField('E_JOURNAL').AsString+'"' ;
  PieceParPiece(Where2,Nomfichier) ;
  QQ.Next ;
  Entet:=False ;
  MoveCur(FALSE) ;
  end ;
ExecuteSql('UPDATE ECRITURE SET E_EXPORTE="X" '+Where) ;
FiniMove ;
end ;

procedure TOF_ExpVersS3.PieceParPiece(Where,NomFichier : string) ;
var QQ : TQuery ;  TheTob : TOB ; i : integer ;
begin
QQ:=OpenSql('SELECT E_ETATLETTRAGE,E_GENERAL,E_IO,E_JOURNAL,E_LETTRAGE,E_NUMEROPIECE,E_REFEXTERNE,'+
'E_TABLE0 AS E_TABLELIBRE1,E_DEBIT,E_ETAT,E_ETATREVISION,E_NIVEAURELANCE,E_REFPOINTAGE,E_UTILISATEUR,'+
'E_TABLE1 AS E_TABLELIBRE2,E_DATEMODIF,E_DATECREATION,E_EXERCICE,E_PAQUETREVISION,E_TAUXDEV,'+
'E_DEVISE,E_COUVERTURE,E_NUMLIGNE,E_DATEECHEANCE,E_DATECOMPTABLE,E_CREDIT,E_ANA,E_DEBITDEV,'+
'E_MODEPAIE AS E_MODEP,E_TABLE2 AS E_TABLELIBRE3,E_QUALIFPIECE,E_AUXILIAIRE,E_LIBELLE,'+
'E_CREDITDEV,E_COTATION'+
' FROM ECRITURE '+Where+' ORDER BY E_GENERAL',True) ;
TheTob:=TOB.Create('ECRITURE_S',nil,-1);
TheTob.LoadDetailDB('_ECRITURE','','',QQ,False);
For i:=0 to TheTob.Detail.Count-1 do FilleParFille(TheTob.Detail[i],NomFichier);
TheTob.Detail[0].DelChampSup('E_COTATION',True);
TheTob.SaveToFile(NomFichier,True,True,True) ;
TheTob.free ;
Ferme(QQ);
end ;

procedure TOF_ExpVersS3.FilleParFille(TheGirl : TOB;NomFichier : string) ;
var TauxDev : double ;Devise,Saisie : string ;
begin
TauxDev:=0 ;
Devise:=TheGirl.GetValue('E_DEVISE') ;
Saisie:='-';
if Devise=V_PGI.DevisePivot then
  begin
    Devise:=DevisePivot ;
  end
  else
  begin
    if not EstMonnaieIn(Devise) Then
      if VH^.TenueEuro then TauxDev:=TheGirl.GetValue('E_COTATION')
                       else TauxDev:=TheGirl.GetValue('E_TAUXDEV');
  end ;
TheGirl.PutValue('E_DEVISE',Devise) ;
TheGirl.PutValue('E_TAUXDEV',TauxDev) ;
RecupAnalytiq(TheGirl,TauxDev,Devise) ;
AjoutChamp(TheGirl) ;
end;

procedure TOF_ExpVersS3.AjoutChamp(TheGirl : TOB) ;
begin
TheGirl.AddChampSup('E_NUMBOR',True) ;   // numero de bordereau de remise en banque ??
TheGirl.PutValue('E_NUMBOR','') ;
end ;

procedure TOF_ExpVersS3.RecupAnalytiq(TheGirl : TOB;TauxDev : double;Devise : string ) ;
var Q2 : TQuery; Sql,Where : string ; CAxes : THValComboBox; i : integer ;
begin
CAxes:=THValComboBox(GetControl('CAXES')) ;
Sql:='SELECT Y_DATECOMPTABLE,Y_DEVISE,Y_LIBELLE,Y_CREDITDEV,Y_NUMLIGNE,Y_JOURNAL,Y_NUMVENTIL,Y_DEBIT,' ;
Sql:=Sql+'Y_REFINTERNE AS Y_REFERENCE,Y_SECTION,Y_TAUXDEV,Y_TABLE2 AS Y_TABLELIBRE3,' ;
Sql:=Sql+'Y_TABLE1 AS Y_TABLELIBRE2,Y_TABLE0 AS Y_TABLELIBRE1,Y_GENERAL,Y_NUMEROPIECE,Y_CREDIT,' ;
Sql:=Sql+'Y_DEBITDEV FROM ANALYTIQ ' ;
Where:=Where+'WHERE Y_DATECOMPTABLE ="'+UsDateTime(TheGirl.GetValue('E_DATECOMPTABLE'))+'" AND ' ;
Where:=Where+'Y_JOURNAL ="'+TheGirl.GetValue('E_JOURNAL')+'" AND ' ;
Where:=Where+'Y_EXERCICE ="'+TheGirl.GetValue('E_EXERCICE')+'" AND ' ;
Where:=Where+'Y_QUALIFPIECE ="'+TheGirl.GetValue('E_QUALIFPIECE')+'" AND ' ;
Where:=Where+'Y_NUMEROPIECE ="'+IntToStr(TheGirl.GetValue('E_NUMEROPIECE'))+'" AND ' ;
Where:=Where+'Y_NUMLIGNE ="'+IntToStr(TheGirl.GetValue('E_NUMLIGNE'))+'" AND ' ;
Where:=Where+'Y_AXE ="'+CAxes.Value+'"' ;
Sql:=Sql+Where ;
Q2:=Opensql(Sql,True) ;
TheGirl.LoadDetailDB('_ANALITIQ','','',Q2,False) ;
for i:=0 to TheGirl.Detail.Count-1 do
  begin
  TheGirl.Detail[i].PutValue('Y_TAUXDEV',TauxDev) ;
  TheGirl.Detail[i].PutValue('Y_DEVISE',Devise) ;
  end ;
ferme(Q2) ;
ExecuteSql('UPDATE ANALYTIQ SET Y_EXPORTE="X" '+Where) ;
end;
{RecupArguments Récupération des arguments de la fiche }
function TOF_ExpVersS3.RecupArguments(var TCritere : TStringList;var NomFichier : string) : boolean;
var CDateDeb,CDatefin,NumPiece1,NumPiece2,SaufJnl,FileName,MPivot,MAutre : THEdit ;FEtab,FNatJnl1,FNatJnl2 : THValComboBox;
Fexport : TCheckBox ; ListeJal,Jal : string ;
begin
Result:=False ;
MPivot:=THEdit(GetControl('MPIVOT')) ;
MAutre:=THEdit(GetControl('MAUTRE')) ;
CDateDeb:=THEdit(GetControl('FDATEECR1')) ;
CDateFin:=THEdit(GetControl('FDATEECR2')) ;
FEtab:=THValComboBox(GetControl('FETAB')) ;
FNatJnl1:=THValComboBox(GetControl('FNATJNL1')) ;
FNatJnl2:=THValComboBox(GetControl('FNATJNL2')) ;
SaufJnl:=THEdit(GetControl('SAUFJAL')) ;
NumPiece1:=THEdit(GetControl('NUMPIECE1')) ;
NumPiece2:=THEdit(GetControl('NUMPIECE2')) ;
FExport:=TCheckBox(GetControl('EXPORT')) ;
FileName:=THEdit(GetControl('FILENAME')) ;
if (CDateDeb=nil) or (CDateFin=nil) or (FEtab=nil) or (FNatJnl1=nil) or (FNatJnl2=nil) or (MPivot=nil)
   or (SaufJnl=nil) or (NumPiece1=nil) or (NumPiece2=nil) or (Fexport=nil) or (FileName=nil) or (Mautre=nil)
   then begin HShowMessage(TMsg[1],'','') ; Exit ;end ;
if ArrangeDate(CDateDeb.Text)>ArrangeDate(CDateFin.Text) then  Begin HShowMessage(TMsg[2],'','') ;Exit ;end ;
if FNatJnl1.ItemIndex>FNatJnl2.ItemIndex then  Begin  HShowMessage(TMsg[3],'','') ;  Exit ;  end ;
if NumPiece1.Text>NumPiece2.Text then  Begin  HShowMessage(TMsg[4],'','') ;  Exit ;  end ;
if MPivot.Text=MAutre.Text then Begin  HShowMessage(TMsg[5],'','') ;  Exit ;  end ;
if (MPivot.Text='') or (MAutre.Text='') then Begin  HShowMessage(TMsg[6],'','') ;  Exit ;  end ;
TCritere.Add('E_DATECOMPTABLE>="'+UsDateTime(StrToDate(CDateDeb.Text))+'"');
TCritere.Add('E_DATECOMPTABLE<="'+UsDateTime(StrToDate(CDateFin.Text))+'"');
TCritere.Add('E_JOURNAL>="'+FNatJnl1.Value+'"');
TCritere.Add('E_JOURNAL<="'+FNatJnl2.Value+'"');
TCritere.Add('E_NUMEROPIECE>="'+NumPiece1.Text+'"');
TCritere.Add('E_NUMEROPIECE<="'+NumPiece2.Text+'"');
if Not FExport.Checked then TCritere.Add('E_EXPORTE<>"X"');
ListeJal:=SaufJnl.Text ;
Jal:=ReadTokenSt(ListeJal) ;
while Jal<>'' do
  begin
  TCritere.Add('E_JOURNAL<>"'+Jal+'"');
  Jal:=ReadTokenSt(ListeJal) ;
  end ;
NomFichier:=FileName.Text ;
DevisePivot:=MPivot.Text ;
DeviseAutre:=MAutre.Text ;
Result:=True ;
end ;
{ RecupTiers récupération des Comptes }
procedure TOF_ExpVersS3.RecupCompte(NomFichier : string) ;
var QQ,Q1 : TQuery ; TheTob : TOB ;
begin
if not TRadioButton(GetControl('CCOMPTE')).Checked then exit ;
Q1:=OpenSql('SELECT G_GENERAL FROM GENERAUX',True) ;
While not Q1.Eof do
  begin
  QQ:=OpenSql('SELECT G_GENERAL,G_LIBELLE,G_FERME,G_NATUREGENE,G_UTILISATEUR,G_PURGEABLE,G_CENTRALISABLE,'
             +'G_DATECREATION,G_DATEMODIF,G_TOTALDEBIT,G_TOTALCREDIT,G_VENTILABLE,G_LETTRABLE,G_DERNLETTRAGE'
             +' FROM GENERAUX WHERE G_GENERAL="'+Q1.FindField('G_GENERAL').AsString+'"',True) ;
  TheTob:=TOB.Create('_GENERAUX',nil,-1);
  TheTob.SelectDB('',QQ,False);
  ferme(QQ) ;
  Q1.Next ;
  TheTob.SaveToFile(NomFichier,True,True,True) ;
  TheTob.free ;
  end ;
  ferme(Q1) ;
end ;
{ RecupTiers récupération des tiers }
procedure TOF_ExpVersS3.RecupTiers(NomFichier : string) ;
var QQ,Q1 : TQuery ; TheTob : TOB ;
begin
if not TRadioButton(GetControl('CTIER')).Checked then exit ;
Q1:=OpenSql('SELECT T_AUXILIAIRE FROM TIERS',True) ;
While not Q1.Eof do
  begin
  QQ:=OpenSql('SELECT T_AUXILIAIRE,T_NATUREAUXI,T_LIBELLE,T_ADRESSE1,T_ADRESSE2,T_CODEPOSTAL,T_VILLE,'
             +'T_PAYS,T_LANGUE,T_DEVISE,T_SIRET,T_APE,T_TELEPHONE,T_FAX,T_REMISE,T_MODEREGLE,T_REGIMETVA,'
             +'T_SOUMISTPF,T_DATECREATION,T_DATEMODIF,T_CONFIDENTIEL,T_UTILISATEUR,T_RVA,'
             +'T_DERNLETTRAGE,T_CREDITACCORDE,T_ESCOMPTE,T_JURIDIQUE AS T_FORMEJURIDIQUE,'
             +'T_TABLE0 AS T_TABLELIBRE1,T_TABLE1 AS T_TABLELIBRE2,T_TABLE2 AS T_TABLELIBRE3,'
             +'T_TOTALCREDIT,T_TOTALDEBIT,T_SECTEUR,T_FERME,T_COLLECTIF,T_PRENOM FROM TIERS '
             +'WHERE T_AUXILIAIRE="'+Q1.FindField('T_AUXILIAIRE').AsString+'"',True) ;
  TheTob:=TOB.Create('_TIERS',nil,-1);
  TheTob.SelectDB('',QQ,False);
  ferme(QQ) ;
  Q1.Next ;
  TheTob.SaveToFile(NomFichier,True,True,True) ;
  TheTob.free ;
  end ;
  ferme(Q1) ;
end;
{ RecupTiers récupération des Sections Analytiques }
procedure TOF_ExpVersS3.RecupSec(NomFichier : string);
var QQ,Q1 : TQuery ; TheTob : TOB ; CAxes : THValComboBox;
begin
if not TRadioButton(GetControl('CSECTION')).Checked then exit ;
CAxes:=THValComboBox(GetControl('CAXES')) ;
Q1:=OpenSql('SELECT S_SECTION FROM SECTION WHERE S_AXE="'+CAxes.Value+'"',True) ;
While not Q1.Eof do
  begin
  QQ:=OpenSql('SELECT S_SECTION,S_LIBELLE,S_DATECREATION,S_DATEMODIF,S_UTILISATEUR,'
             +'S_TOTALCREDIT,S_TOTALDEBIT FROM SECTION'
             +' WHERE S_AXE="'+CAxes.Value+'" AND S_SECTION="'+Q1.FindField('S_SECTION').AsString+'"',True) ;
  TheTob:=TOB.Create('_SECTION',nil,-1);
  TheTob.SelectDB('',QQ,False);
  ferme(QQ) ;
  Q1.Next ;
  TheTob.SaveToFile(NomFichier,True,True,True) ;
  TheTob.free ;
  end ;
  ferme(Q1) ;
end ;
{ RecupTiers récupération des Journaux }
procedure TOF_ExpVersS3.RecupJal(NomFichier : string);
var QQ,Q1 : TQuery ; TheTob : TOB ;
begin
if not TRadioButton(GetControl('CJOURNAUX')).Checked then exit ;
Q1:=OpenSql('SELECT J_JOURNAL FROM JOURNAL',True) ;
While not Q1.Eof do
  begin
  QQ:=OpenSql('SELECT J_JOURNAL,J_LIBELLE,J_NATUREJAL,J_DATECREATION,J_DATEMODIF,J_UTILISATEUR,'
             +'J_CONTREPARTIE,J_COMPTEINTERDIT FROM JOURNAL'
             +' WHERE J_JOURNAL="'+Q1.FindField('J_JOURNAL').AsString+'"',True) ;
  TheTob:=TOB.Create('_JOURNAL',nil,-1);
  TheTob.SelectDB('',QQ,False);
  ferme(QQ) ;
  Q1.Next ;
  TheTob.SaveToFile(NomFichier,True,True,True) ;
  TheTob.free ;
  end ;
  ferme(Q1) ;
end ;
{ RecupTiers récupération des Journaux }
procedure TOF_ExpVersS3.RecupDevise(NomFichier : string);
var QQ,Q1 : TQuery ; TheTob : TOB ; Devise : string ;
begin
if not TRadioButton(GetControl('CDEVISE')).Checked then exit ;
Q1:=OpenSql('SELECT D_DEVISE,D_MONNAIEIN FROM DEVISE WHERE D_FONGIBLE<>"X"',True) ;
While not Q1.Eof do
  begin
  Devise:=Q1.FindField('D_DEVISE').AsString ;
  if Devise<>V_PGI.DevisePivot then
    begin
    if Q1.FindField('D_MONNAIEIN').AsString='X' then
       QQ:=OpenSql('SELECT D_DEVISE,D_LIBELLE,D_PARITEEURO AS D_COURS,D_DECIMALE AS D_NBDEC,'
               +'D_MONNAIEIN AS D_INEURO FROM DEVISE WHERE D_DEVISE="'+Devise+'" ',True)
       else
       QQ:=OpenSql('SELECT D_DEVISE,D_LIBELLE,H_COTATION AS D_COURS,D_DECIMALE AS D_NBDEC,'
               +'D_MONNAIEIN AS D_INEURO FROM DEVISE LEFT JOIN CHANCELL ON D_DEVISE=H_DEVISE '
               +'WHERE D_DEVISE="'+Devise+'" '
               +'AND H_DATECOURS=(SELECT MAX(H_DATECOURS) FROM CHANCELL WHERE H_DEVISE="'+Devise+'")',True) ;
    TheTob:=TOB.Create('_DEVISE',nil,-1);
    TheTob.SelectDB('',QQ,False);
    ferme(QQ) ;
    TheTob.SaveToFile(NomFichier,True,True,True) ;
    TheTob.free ;
    end ;
    Q1.Next ;
  end ;
  ferme(Q1) ;
end ;
{+===================================+}
{+                                   +}
{+         TOF : EXPVERSS3_2         +}
{+                                   +}
{+===================================+}
procedure TOF_ExpVersS3_2.OnLoad;
var GJal : THGrid ; LItem,LValu : TStringList ; i : integer ;
begin
GJal:=THGrid(GetControl('GJAL')) ;
GJal.ColWidths[SP_COCHE]:=0;
GJal.ColWidths[SP_VALJAL]:=0;
GJal.MultiSelect:=True ;
LItem:=TStringList.Create ;
LValu:=TStringList.Create ;
RemplirValCombo('TTJOURNAL','','',LItem,LValu,False,False) ;
For i:=0 to LItem.Count-1 do
  begin
  GJal.Cells[SP_NOMJAL,i]:=LItem[i] ;
  GJal.Cells[SP_VALJAL,i]:=LValu[i] ;
  GJal.Cells[SP_COCHE,i]:='' ;
  end ;
GJal.RowCount:=LItem.Count ;
end ;

procedure TOF_ExpVersS3_2.OnUpdate;
var GJal : THGrid ; i : integer ; SaufJal : String ;
begin
GJal:=THGrid(GetControl('GJAL')) ;
SaufJal:='' ;
For i:=1 to GJal.RowCount-1 do
  if GJal.IsSelected(i) then SaufJal:=SaufJal+GJal.Cells[SP_VALJAL,i]+';' ;
TFVierge(Ecran).Retour:=SaufJal ;
end ;

Initialization
registerclasses([TOF_ExpVersS3]);
registerclasses([TOF_ExpVersS3_2]);
end.
