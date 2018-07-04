unit ZEcrRegul;

interface

uses
 Classes, SysUtils,
 {$IFNDEF EAGLCLIENT}
 db,dbTables,
 {$ENDIF}
 Ent1,
 HEnt1,
 HCtrls,
 UtilSais,
 ULibEcriture, // pour le BlocageBor
 UTOB;

type

 TEcrRegul= Class
  private
   FTOBEcrGene     : TOB;   // TOB des ecritures de regul pour le mode PCL
   FZListJournal   : TZListJournal; // objet de gestion des journaux
   procedure NumeroLigne ( vTOBPiece : TOB );
   function SaveBordereau ( vTOBEcr : TOB): boolean;
   function SavePiece ( vTOBEcr : TOB): boolean;
    procedure RecupereLigneBor(vTOB: TOB);
    procedure RecupereLignePiece(vTOB: TOB);
  public
   constructor Create;
   destructor  Destroy; override;

   procedure Save;
   procedure RecupereLigne( vTOB : TOB );
   procedure Delete( vStCodeLettrage : string ) ;

   property TOBEcrGene : TOB read FTOBEcrGene write FTOBEcrGene;

 end;


implementation

constructor TEcrRegul.Create;
begin
 FTOBEcrGene   := TOB.Create('',nil,-1);
 FZListJournal := TZListJournal.Create; // Objet de gestion des journaux
end;

destructor TEcrRegul.Destroy;
begin
 FTOBEcrGene.Free;
 FZListJournal.Free;
end;

procedure TEcrRegul.RecupereLigneBor( vTOB : TOB );
var
 lTOBPiece : TOB;
 lStCrit : string;
 lInNumGroupeEcr : integer;
begin
 lStCrit:=vTOB.Detail[0].GetValue('E_JOURNAL')+vTOB.Detail[0].GetValue('E_MODESAISIE')+
          intToStr(vTOB.Detail[0].GetValue('E_PERIODE'))+intToStr(vTOB.Detail[0].GetValue('E_NUMEROPIECE')) ;
 // recherche s'il n'existe pas une piece avec ces critères
 lTOBPiece:=FTOBEcrGene.FindFirst(['CRIT'],[lStCrit],true) ;
 if lTOBPiece=nil then
  BEGIN
   lTOBPiece:=TOB.Create('',FTOBEcrGene,-1) ;
   lTOBPiece.AddChampSup('CRIT',true) ; lTOBPiece.PutValue('CRIT',lStCrit) ;
   lTOBPiece.AddChampSup('LETTRAGE',true) ; lTOBPiece.PutValue('LETTRAGE',vTOB.Detail[vTOB.Detail.count-1].GetValue('E_LETTRAGE')) ;
   lTOBPiece.AddChampSup('MODESAISIE',true) ; lTOBPiece.PutValue('MODESAISIE',vTOB.Detail[vTOB.Detail.count-1].GetValue('E_MODESAISIE')) ;
  END; // if
 if vTOB.Detail[0].GetValue('E_MODESAISIE')<>'BOR' then lInNumGroupeEcr:=1
 else if (lTOBPiece.Detail.Count=0) then lInNumGroupeEcr:=1
 else lInNumGroupeEcr:=lTOBPiece.Detail[lTOBPiece.Detail.Count-1].GetValue('E_NUMGROUPEECR')+1;
 while (vTOB.Detail.Count>0) do
  BEGIN
   vTOB.Detail[0].PutValue('E_NUMGROUPEECR',lInNumGroupeEcr) ; vTOB.Detail[0].ChangeParent(lTOBPiece, -1) ;
  END; // while
end;

procedure TEcrRegul.RecupereLignePiece( vTOB : TOB );
var
 lTOBPiece : TOB;
begin
 lTOBPiece:=TOB.Create('',FTOBEcrGene,-1) ;
 lTOBPiece.AddChampSup('CRIT',true) ; lTOBPiece.PutValue('CRIT','') ;
 lTOBPiece.AddChampSup('LETTRAGE',true) ; lTOBPiece.PutValue('LETTRAGE',vTOB.Detail[vTOB.Detail.count-1].GetValue('E_LETTRAGE')) ;
 lTOBPiece.AddChampSup('MODESAISIE',true) ; lTOBPiece.PutValue('MODESAISIE',vTOB.Detail[vTOB.Detail.count-1].GetValue('E_MODESAISIE')) ;
 while ( vTOB.Detail.Count > 0 ) do
  BEGIN
   vTOB.Detail[0].PutValue('E_NUMGROUPEECR', 1 ) ; vTOB.Detail[0].ChangeParent(lTOBPiece, -1) ;
  END; // while
end;


procedure TEcrRegul.RecupereLigne( vTOB : TOB );
begin
 if vTOB=nil then exit ; if vTOB.Detail.Count=0 then exit ;

 if vTOB.Detail[0].GetValue('E_MODESAISIE') = '-' then
  RecupereLignePiece( vTOB )
   else
    RecupereLigneBor( vTOB );

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 25/04/2002
Modifié le ... :   /  /
Description .. : numerotation des lignes du bordereau et des lignes
Suite ........ : d'analytique
Mots clefs ... :
*****************************************************************}
procedure TEcrRegul.NumeroLigne(vTOBPiece : TOB );
var
 lInNumGroupEcr     : integer;
 lInNumLigne        : integer;
 k                  : integer;
 i                  : integer;
 lNumAxe            : integer;
 lInNumLignePrec    : integer;
 lQ                 : TQuery;
 lTOBLigneEcr       : TOB;
 lTOBSection        : TOB;
begin
 lInNumLignePrec:=0 ;
 // on recherche le dernier numero  de ligne et de numgroupecr
 lQ:=OpenSQL('select max(E_NUMLIGNE) as N ,max(E_NUMGROUPEECR) as M from ecriture where E_EXERCICE="'+vTOBPiece.Detail[0].GetValue('E_EXERCICE')+'" '+
             'and E_NUMEROPIECE='+intToStr(vTOBPiece.Detail[0].GetValue('E_NUMEROPIECE'))+
             'and E_JOURNAL="'+vTOBPiece.Detail[0].GetValue('E_JOURNAL')+'" and E_PERIODE='+intToStr(vTOBPiece.Detail[0].GetValue('E_PERIODE')),true);
 lInNumLigne:=lQ.FindField('N').asInteger ; Inc(lInNumLigne) ; lInNumGroupEcr:=lQ.FindField('M').asInteger ; Ferme(lQ) ;
 // on parcourt l'ensemble des lignes de la piece et on numerote les lignes
 for k:=0 to vTOBPiece.Detail.Count-1 do
  BEGIN
   lTOBLigneEcr := vTOBPiece.Detail[k];
   lTOBLigneEcr.PutValue('E_NUMLIGNE',lInNumLigne) ;
   // renumerotation de l'analytique
   if assigned( lTOBLigneEcr.Detail ) and ( lTOBLigneEcr.detail.Count > 0 ) then
    begin
     for lNumAxe := 0 to 4 do
      begin
       lTOBSection := lTOBLigneEcr.detail[lNumAxe];
       if assigned( lTOBSection.Detail ) and ( lTOBSection.detail.Count > 0 ) then
         for i := 0 to lTOBSection.detail.Count - 1 do
          lTOBSection.Detail[i].PutValue('Y_NUMLIGNE',lInNumLigne);
      end; // for
    end; // if
   Inc(lInNumLigne) ;
   if lTOBLigneEcr.GetValue('E_MODESAISIE')='BOR' then
    BEGIN
     if lTOBLigneEcr.GetValue('E_NUMGROUPEECR')<>lInNumLignePrec then
      BEGIN
       lInNumLignePrec:=lTOBLigneEcr.GetValue('E_NUMGROUPEECR') ; Inc(lInNumGroupEcr) ;
      END;
    END // if
     else lInNumGroupEcr:=1;
   // on supprime les eventuelles caractères speciaux du lettrages
   if lTOBLigneEcr.GetValue('E_LETTRAGE')<>'' then
    lTOBLigneEcr.PutValue('E_LETTRAGE',Copy(vTOBPiece.Detail[k].GetValue('E_LETTRAGE'),1,4));
   lTOBLigneEcr.PutValue('E_NUMGROUPEECR',lInNumGroupEcr);
  END;
end;


function TEcrRegul.SaveBordereau( vTOBEcr : TOB ) : boolean;
begin

 result := false;

 if BlocageBor( vTOBEcr.Detail[0].GetValue('E_EXERCICE'),
                vTOBEcr.Detail[0].GetValue('E_JOURNAL'),
                vTOBEcr.Detail[0].GetValue('E_DATECOMPTABLE'),
                vTOBEcr.Detail[0].GetValue('E_NUMEROPIECE'),true) then
    begin
     try

      NumeroLigne(vTOBEcr) ;
      vTOBEcr.InsertDBByNivel(false);
      MajSoldesEcritureTOB (vTOBEcr,false);

      result := true ;

     finally
      BloqueurBor( vTOBEcr.Detail[0].GetValue('E_EXERCICE'),
                   vTOBEcr.Detail[0].GetValue('E_JOURNAL'),
                   vTOBEcr.Detail[0].GetValue('E_DATECOMPTABLE'),
                   vTOBEcr.Detail[0].GetValue('E_NUMEROPIECE'), true);
     end; //try

    end
     else
      begin
       result := false;
       CAfficheLigneEcrEnErreur(vTOBEcr.Detail[0]) ;
       V_PGI.IOError:=oeStock ;
      end ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 06/06/2002
Modifié le ... :   /  /    
Description .. : -06/06/2002- on affecte le numero de piece au moment de 
Suite ........ : l'enregistrement
Mots clefs ... : 
*****************************************************************}
function TEcrRegul.SavePiece( vTOBEcr : TOB ) : boolean;
var
 lInNum       : integer;
 lNumAxe      : integer;
 i,k          : integer;
 lTOBLigneEcr : TOB;
 lTOBSection  : TOB;
begin
 result:=false;
 
 FZListJournal.Load([vTOBEcr.Detail[0].GetValue('E_JOURNAL')]);
 lInNum  := FZListJournal.GetNumJal( StrToDate(vTOBEcr.Detail[0].GetValue('E_DATECOMPTABLE')));

 for k:=0 to vTOBEcr.Detail.Count - 1 do
  begin
   lTOBLigneEcr := vTOBEcr.Detail[k];
   lTOBLigneEcr.PutValue('E_NUMEROPIECE',lInNum);
   // renumerotation de l'analytique
   if assigned( lTOBLigneEcr.Detail ) and ( lTOBLigneEcr.detail.Count > 0 ) then
    begin
     for lNumAxe := 0 to 4 do
      begin
       lTOBSection := lTOBLigneEcr.detail[lNumAxe];
       if assigned( lTOBSection.Detail ) and ( lTOBSection.detail.Count > 0 ) then
         for i := 0 to lTOBSection.detail.Count - 1 do
          begin
           lTOBSection.Detail[i].PutValue('Y_NUMLIGNE'   , lTOBLigneEcr.GetValue('E_NUMLIGNE'));
           lTOBSection.Detail[i].PutValue('Y_NUMEROPIECE', lInNum);
          end; // for
      end; // for
    end; // if


  end; // for

 vTOBEcr.InsertDBByNivel(false);
 MajSoldesEcritureTOB (vTOBEcr,false);
 result:= true;
end;

procedure TEcrRegul.Save;
var
 i         : integer ;
 lTOBPiece : TOB;
begin
 for i:=0 to ( FTOBEcrGene.Detail.Count - 1 ) do
  begin
   lTOBPiece := FTOBEcrGene.Detail[i];
   if ( lTOBPiece.Detail[0].GetValue('E_MODESAISIE')='BOR' ) or ( lTOBPiece.Detail[0].GetValue('E_MODESAISIE')='LIB' ) then
    begin
     if not SaveBordereau( lTOBPiece ) then break;
    end
     else
      if not SavePiece( lTOBPiece ) then break;
  end; // for
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 17/04/2002
Modifié le ... :   /  /
Description .. : Supprime une piece suite à la suppression du lettrage
Suite ........ : associée
Mots clefs ... :
*****************************************************************}
procedure TEcrRegul.Delete( vStCodeLettrage : string );
var
 lInIndex : integer;
begin
 lInIndex := 0;

 while lInIndex <= ( FTOBEcrGene.Detail.Count-1 ) do
  begin
   if FTOBEcrGene.Detail[lInIndex].GetValue('LETTRAGE')= Copy(vStCodeLettrage,1,4) then
     FTOBEcrGene.Detail[lInIndex].Free;
   Inc(lInIndex);
  end; // while
end;

end.
