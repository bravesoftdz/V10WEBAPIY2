{***********UNITE*************************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 25/11/2002
Modifié le ... :   /  /    
Description .. : Unit qui permet de re générer la fatcuration éclatée par 
Suite ........ : assistant sur une période de date.
Suite ........ : le fait pour toutes les missions qui ont de l'activité sur la 
Suite ........ : période sélectionnée et pour toutes les factures de ces 
Suite ........ : missions
Mots clefs ... : 
*****************************************************************}
Unit UTOFAFREGEN_FACTECLAT ;

Interface

Uses StdCtrls,  Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db, 
     dbtables, FE_Main ,
{$ELSE}
    Maineagl,
{$ENDIF}
   Utof, hmsgbox, dicoAF, Utob, Hctrls ,SysUtils ,Entgc ,FactUtil ,
   AffaireUtil, FactComm, FactTOB
   ;

Type
  TOF_AFREGEN_FACTECLAT= Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnUpdate                 ; override ;
    Procedure BoucleRegenEclatFact;
  Private
    Erreur:Boolean;
    end ;
Procedure AFLanceFiche_REgen_FActEclat;

Implementation


procedure TOF_AFREGEN_FACTECLAT.OnArgument (S : String ) ;
begin
  Inherited ;
Erreur:=False;
end ;

procedure TOF_AFREGEN_FACTECLAT.OnClose ;

begin
  Inherited ;
If Erreur then begin
   LastError:=(-1);
   LastErrorMsg:='';
   Erreur:=false;
   exit;     // il y eu une erreur, on reste sur la même fiche
   end;
end ;

procedure TOF_AFREGEN_FACTECLAT.OnUpdate;
var text,titre:string;
begin
  Inherited ;
  Titre :='Regénération facture éclatée';
  text:='Confirmez-vous la regénération de la facturation éclatée par ressource sur cette période';
  If PGIAsk(traduitGA(text),titre)<> mrYes then exit;
  SetControlText('SUIVI','Début traitement');
  BoucleRegenEclatFact;
  SetControlText('SUIVI','Traitement terminé');
end ;

procedure TOF_AFREGEN_FACTECLAT.BoucleRegenEclatFact ;
var   qq:Tquery;
 TobAff ,Tobdet, TobGLo, TobPiece,TobPieceDet, TObDetaff : Tob;
 Sql : string;
 CleDoc : R_CleDoc ;
 ii, wi ,wj: integer;
begin
  TobAff := Tob.create ('les aff',Nil, -1) ; // mcd 28/01/03 chgmt nom liste
  Sql :=  'SELECT ACT_AFFAIRE FROM ACTIVITE WHERE ACT_DATEACTIVITE >="'
      + UsDateTime(StrTodate(GetControlText('DATEDEB')))+'" AND ACT_DATEACTIVITE<="'
      + UsDAteTIme(StrTodate(GetControlText('DATEFIN'))) +'" AND ACT_TYPEACTIVITE="REA" GROUP BY ACT_AFFAIRE';
  QQ := OpenSql (Sql , true);
  if Not QQ.EOF then TobAff.loadDetailDb( 'detail','','',QQ,False);
  Ferme(QQ);
  for ii:=0 to TobAff.detail.count -1 do
    begin
    TobdetAff := TobAff.detail[ii];
    SetControlText('SUIVI',Format (TraduitGa('Affaire en cours %s '), [TobdetAff.getValue('ACT_AFFAIRE')]));
    TobGlo :=Tob.Create('Liste piece',Nil,-1);
    QQ:=OpenSql('SELECT GL_NATUREPIECEG,Gl_DATEPIECE,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_AFFAIRE FROM LIGNE WHERE (GL_NATUREPIECEG="FAC" OR GL_NATUREPIECEG="FRE" OR GL_NATUREPIECEG="AVC") '
      + ' AND GL_AFFAIRE ="' + TobDetAff.getvalue('ACT_AFFAIRE')
      + '" AND GL_DATEPIECE >="' +UsDateTime(StrTodate(GetControlText('DATEDEBFAC')))
      + '" AND GL_DATEPIECE <="' +UsDateTime(StrTodate(GetControlText('DATEFIN'))) + '"'
      + 'GROUP BY GL_NATUREPIECEG,GL_DATEPIECE,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_AFFAIRE',True);
    if Not QQ.EOF then TobGlo.loadDetailDb( 'Detail','','',QQ,False);
    Ferme(QQ);
    if TObGlo.detail.count >=1 then ExecuteSql ('DELETE AFCUMUL WHERE ACU_TYPEAC="FAC" AND ACU_AFFAIRE="'+
      +  TobDetAff.getvalue('ACT_AFFAIRE')
      + '" AND ACU_DATE >="' +UsDateTime(StrTodate(GetControlText('DATEDEBFAC')))
      + '" AND ACU_DATE <="' +UsDateTime(StrTodate(GetControlText('DATEFIN'))) + '"');
    for wi := 0 to TObGlo.detail.count-1 do
      begin
      tobDet:=TobGlo.detail[wi];
      CleDoc.NaturePiece:=TobDEt.getValue('GL_NATUREPIECEG') ;
      CleDoc.DatePiece:=TobDEt.getValue('GL_DATEPIECE') ;
      CleDoc.Souche:=TobDEt.getValue('GL_SOUCHE') ;
      CleDoc.NumeroPiece:=TobDEt.getValue('GL_NUMERO') ;
      CleDoc.Indice:=TobDEt.getValue('GL_INDICEG');
      TOBPiece:=TOB.Create('PIECE',Nil,-1) ;
         // charge une seule facture.. on peut se permettre de tout prendre
      QQ:=OpenSQL('SELECT * FROM PIECE WHERE '+WherePiece(CleDoc,ttdPiece,False),True) ;
      TOBPiece.SelectDB('',QQ) ;
      Ferme(QQ) ;
      // Lecture Lignes
         // charge une seule facture.. on peut se permettre de tout prendre  mcd 14/01/03 sur la même aff seulement
      QQ:=OpenSQL('SELECT * FROM LIGNE WHERE '+WherePiece(CleDoc,ttdLigne,False)
        + ' AND GL_AFFAIRE="'+TobDet.getValue('GL_AFFAIRE')+'" ORDER BY GL_NUMLIGNE',True) ;
      TOBPiece.LoadDetailDB('LIGNE','','',QQ,False,True) ;
      Ferme(QQ) ;
      PieceAjouteSousDetail(TOBPiece);
      MajFactEclat(TobPiece);
      TobPiece.free;
      end;
    TobGLo.free;
    end;
  tobaff.free;
end;

Procedure AFLanceFiche_REgen_FActEclat;
begin
if ToutSeulAff  then  exit; // quelqu'un d'autre travaille sur la base
try
AGLLanceFiche ('AFF','AFREGEN_FACTECLAT','','','');
finally
Bloqueur ('AffToutSeul',False);
end;
end;

Initialization
  registerclasses ( [ TOF_AFREGEN_FACTECLAT] ) ;
end.
