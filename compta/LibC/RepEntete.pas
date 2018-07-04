{***********UNITE*************************************************
Auteur  ...... : BPY
Créé le ...... : 15/06/2004
Modifié le ... : 17/06/2004
Description .. : Réparation des entêtes de pieces !!
Suite ........ : pb lier a un clien qui a utilisé la modif entête de piece alors 
Suite ........ : que celle ci etait cassé !
Suite ........ : uniquement pour la saisie pieces de PGE .. ne fonctionne 
Suite ........ : pas sur les bordereau !
Mots clefs ... : MODIF;ENTETE;PIECE;REPARATION
*****************************************************************}

unit RepEntete;

//================================================================================
// Interface
//================================================================================
interface

uses
    Windows,
    {$IFNDEF EAGLCLIENT}
    db,
   {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
   {$ENDIF}
    Messages,
    SysUtils,
    Classes,
    HPanel,
    Graphics,
    UiUtil,
    Controls,
    Forms,
    Dialogs,
    StdCtrls,
    HTB97,
    Hctrls,
    HEnt1,
    UTOB,
    SaisUtil ,
    HMsgBox
    ;

//==================================================
// Externe
//==================================================
Procedure ReparationEntete;

//==================================================
// Definition de class
//==================================================
type
    TFRepEntete = class(TForm)
        Dock: TDock97;
        ToolWindow971: TToolWindow97;
        BAide: TToolbarButton97;
        BValider: TToolbarButton97;
        BClose: TToolbarButton97;
        CBNature: TCheckBox;
        CBEtab: TCheckBox;
        CBDate: TCheckBox;
    CPiece: TCheckBox;

        procedure BValiderClick(Sender: TObject);
        procedure BAideClick(Sender: TObject);
    private
        { Déclarations privées }

       function ReparePiece : boolean ; 

    public
        { Déclarations publiques }
    end;

//================================================================================
// Implementation
//================================================================================
implementation

{$R *.DFM}

//==================================================
// fonctions hors class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 15/06/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure ReparationEntete;
var
    FRepEntete: TFRepEntete;
    PP : THPanel;
begin
    FRepEntete := TFRepEntete.Create(Application);
    PP :=  FindInsidePanel;

    if (PP = nil) then
    begin
        try
            FRepEntete.ShowModal;
        finally
            FRepEntete.Free;
        end;
    end
    else
    begin
        InitInside(FRepEntete,PP);
        FRepEntete.Show;
    end;
end;


function TFRepEntete.ReparePiece : boolean ;
var
 lQ : TQuery ;
 lT1,lT2 : TOB ;
 i,j : integer ;
 lInNum: integer ;
begin

 result := true ;

 try

 lQ := OpenSQl('select e_exercice,e_journal,e_numeropiece from ecriture , exercice ' +
                 ' where e_modesaisie="-" and e_valide="-" and e_exercice=ex_exercice and ex_etatcpta="OUV" and E_QUALIFPIECE="N" ' +
                 ' and e_numligne=1 ' +
                 ' group by e_exercice,e_journal,e_numeropiece ' +
                 ' having count(*) > 1 ' +
                 ' order by e_exercice,e_journal,e_numeropiece ' , true ) ;

 lT1 := TOB.Create('',nil,-1) ;
 lT1.LoadDetailDB('','','',lQ,true) ;
 ferme(lQ) ;

 for i := 0 to lT1.Detail.Count - 1 do
  begin
   lQ := OpenSql('select E_JOURNAL,e_datecomptable,e_exercice,e_numeropiece ' +
                 ' from ecriture ' +
                 ' where e_journal="' + lT1.Detail[i].GetValue('e_journal') + '" ' +
                 'and e_exercice="' + lT1.Detail[i].GetValue('e_exercice') + '" ' +
                 ' and E_QUALIFPIECE="N" ' +
                 'and e_numeropiece=' + intToStr(lT1.Detail[i].GetValue('e_numeropiece'))  +
                 ' and e_numligne = 1' , true ) ;

   lT2 := TOB.Create('',nil,-1) ;
   lT2.LoadDetailDB('','','',lQ,true) ;
   ferme(lQ) ;

   for j := 0 to lT2.Detail.Count - 1 do
    begin

   lInNum := GetNewNumJal ( lT2.Detail[j].GetValue('E_JOURNAL') , true , lT2.Detail[j].GetValue('E_DATECOMPTABLE')) ;

   executeSql('update ecriture set e_io="-",e_numeropiece=' + intToStr(lInNum) +
              ' where e_journal="' + lT2.Detail[j].GetValue('e_journal') + '" ' +
              ' and e_exercice="' + lT2.Detail[j].GetValue('e_exercice') + '" ' +
              ' and e_numeropiece=' + intToStr(lT2.Detail[j].GetValue('e_numeropiece')) +
              '  and e_datecomptable="' + usDatetime(lT2.Detail[j].GetValue('E_DATECOMPTABLE')) + '" ' ) ;

   executeSql('UPDATE ANALYTIQ SET Y_NUMEROPIECE=' + intToStr(lInNum) + ' ' +
             'WHERE Y_JOURNAL="'      + lT2.Detail[j].GetValue('e_journal')    + '" ' +
             'AND Y_EXERCICE="'       + lT2.Detail[j].GetValue('e_exercice')   + '" ' +
             'AND Y_NUMEROPIECE='     + intToStr(lT2.Detail[j].GetValue('e_numeropiece')) + ' ' +
             'AND Y_DATECOMPTABLE="'  + usDatetime(lT2.Detail[j].GetValue('E_DATECOMPTABLE'))  + '" ' ) ;
    end ;

   lT2.Free ;
   
  end ;

 lT1.Free ;

 except
  on E : Exception do
   begin
    result := false ;
    PGIBox(TraduireMemoire('Erreur lors de la réparation.' + #10 + #13 + 'Celle ci a été annulée.')+ #10#13 + E.Message);
    SourisNormale;
    EnableControls(Self,true);
   end ;
  end ;

end ;




//==================================================
// Evenements de la fiche
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 15/06/2004
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TFRepEntete.BValiderClick(Sender: TObject);
var
    SQLE,SQLA,erreur : string;
begin
    EnableControls(Self,false);
    SourisSablier;
    erreur := TraduireMemoire('Erreur lors de la réparation.' + #10 + #13 + 'Celle ci a été annulée.');

    BeginTrans;

    // Nature de pieces
    if (CBNature.Checked) then
    begin
        // ECRITURE
        SQLE := 'UPDATE ECRITURE SET';
        SQLE := SQLE + ' E_NATUREPIECE = (SELECT E1.E_NATUREPIECE FROM ECRITURE E1 WHERE';
        SQLE := SQLE + ' E1.E_JOURNAL = ECRITURE.E_JOURNAL AND E1.E_EXERCICE = ECRITURE.E_EXERCICE AND E1.E_DATECOMPTABLE = ECRITURE.E_DATECOMPTABLE AND E1.E_NUMEROPIECE = ECRITURE.E_NUMEROPIECE AND E1.E_NUMLIGNE';
        SQLE := SQLE + ' = (SELECT MAX(E2.E_NUMLIGNE) FROM ECRITURE E2 WHERE E2.E_JOURNAL = ECRITURE.E_JOURNAL AND E2.E_EXERCICE = ECRITURE.E_EXERCICE AND E2.E_DATECOMPTABLE = ECRITURE.E_DATECOMPTABLE AND ';
        SQLE := SQLE + ' E2.E_NUMEROPIECE = ECRITURE.E_NUMEROPIECE AND E2.E_QUALIFPIECE = ECRITURE.E_QUALIFPIECE) AND E1.E_NUMECHE IN (0,1) AND E1.E_QUALIFPIECE = ECRITURE.E_QUALIFPIECE) WHERE E_MODESAISIE="-"';
        try
            ExecuteSQL(SQLE);
        except
            PGIBox(erreur);
            RollBack;
            SourisNormale;
            EnableControls(Self,true);
            exit;
        end;
        // ANALYTIQUE
        SQLA := 'UPDATE ANALYTIQ SET';
        SQLA := SQLA + ' Y_NATUREPIECE = (SELECT A1.Y_NATUREPIECE FROM ANALYTIQ A1 WHERE';
        SQLA := SQLA + ' A1.Y_JOURNAL = ANALYTIQ.Y_JOURNAL AND A1.Y_EXERCICE = ANALYTIQ.Y_EXERCICE AND A1.Y_DATECOMPTABLE = ANALYTIQ.Y_DATECOMPTABLE AND A1.Y_NUMEROPIECE = ANALYTIQ.Y_NUMEROPIECE AND A1.Y_NUMLIGNE';
        SQLA := SQLA + ' = (SELECT MAX(A2.Y_NUMLIGNE) FROM ANALYTIQ A2 WHERE A2.Y_JOURNAL = ANALYTIQ.Y_JOURNAL AND A2.Y_EXERCICE = ANALYTIQ.Y_EXERCICE AND A2.Y_DATECOMPTABLE = ANALYTIQ.Y_DATECOMPTABLE AND';
        SQLA := SQLA + ' A2.Y_NUMEROPIECE = ANALYTIQ.Y_NUMEROPIECE AND A2.Y_AXE = ANALYTIQ.Y_AXE AND A2.Y_NUMVENTIL = ANALYTIQ.Y_NUMVENTIL AND A2.Y_QUALIFPIECE =  ANALYTIQ.Y_QUALIFPIECE) AND A1.Y_AXE';
        SQLA := SQLA + ' = ANALYTIQ.Y_AXE AND A1.Y_NUMVENTIL = ANALYTIQ.Y_NUMVENTIL AND A1.Y_QUALIFPIECE = ANALYTIQ.Y_QUALIFPIECE) WHERE EXISTS (SELECT E_EXERCICE FROM ECRITURE  WHERE E_JOURNAL = Y_JOURNAL';
        SQLA := SQLA + ' AND E_EXERCICE = Y_EXERCICE AND E_DATECOMPTABLE = Y_DATECOMPTABLE AND E_NUMEROPIECE = Y_NUMEROPIECE AND E_NUMLIGNE = Y_NUMLIGNE AND E_QUALIFPIECE = Y_QUALIFPIECE AND E_MODESAISIE="-")';
        try
            ExecuteSQL(SQLA);
        except
            PGIBox(erreur);
            RollBack;
            SourisNormale;
            EnableControls(Self,true);
            exit;
        end;
    end;

    // Etablissement
    if (CBEtab.Checked) then
    begin
        // ECRITURE
        SQLE := 'UPDATE ECRITURE SET';
        SQLE := SQLE + ' E_ETABLISSEMENT = (SELECT E1.E_ETABLISSEMENT FROM ECRITURE E1 WHERE';
        SQLE := SQLE + ' E1.E_JOURNAL = ECRITURE.E_JOURNAL AND E1.E_EXERCICE = ECRITURE.E_EXERCICE AND E1.E_DATECOMPTABLE = ECRITURE.E_DATECOMPTABLE AND E1.E_NUMEROPIECE = ECRITURE.E_NUMEROPIECE AND E1.E_NUMLIGNE';
        SQLE := SQLE + ' = (SELECT MAX(E2.E_NUMLIGNE) FROM ECRITURE E2 WHERE E2.E_JOURNAL = ECRITURE.E_JOURNAL AND E2.E_EXERCICE = ECRITURE.E_EXERCICE AND E2.E_DATECOMPTABLE = ECRITURE.E_DATECOMPTABLE AND ';
        SQLE := SQLE + ' E2.E_NUMEROPIECE = ECRITURE.E_NUMEROPIECE AND E2.E_QUALIFPIECE = ECRITURE.E_QUALIFPIECE) AND E1.E_NUMECHE IN (0,1) AND E1.E_QUALIFPIECE = ECRITURE.E_QUALIFPIECE) WHERE E_MODESAISIE="-"';
        try
            ExecuteSQL(SQLE);
        except
            PGIBox(erreur);
            RollBack;
            SourisNormale;
            EnableControls(Self,true);
            exit;
        end;
        // ANALYTIQUE
        SQLA := 'UPDATE ANALYTIQ SET';
        SQLA := SQLA + ' Y_ETABLISSEMENT = (SELECT A1.Y_ETABLISSEMENT FROM ANALYTIQ A1 WHERE';
        SQLA := SQLA + ' A1.Y_JOURNAL = ANALYTIQ.Y_JOURNAL AND A1.Y_EXERCICE = ANALYTIQ.Y_EXERCICE AND A1.Y_DATECOMPTABLE = ANALYTIQ.Y_DATECOMPTABLE AND A1.Y_NUMEROPIECE = ANALYTIQ.Y_NUMEROPIECE AND A1.Y_NUMLIGNE';
        SQLA := SQLA + ' = (SELECT MAX(A2.Y_NUMLIGNE) FROM ANALYTIQ A2 WHERE A2.Y_JOURNAL = ANALYTIQ.Y_JOURNAL AND A2.Y_EXERCICE = ANALYTIQ.Y_EXERCICE AND A2.Y_DATECOMPTABLE = ANALYTIQ.Y_DATECOMPTABLE AND';
        SQLA := SQLA + ' A2.Y_NUMEROPIECE = ANALYTIQ.Y_NUMEROPIECE AND A2.Y_AXE = ANALYTIQ.Y_AXE AND A2.Y_NUMVENTIL = ANALYTIQ.Y_NUMVENTIL AND A2.Y_QUALIFPIECE =  ANALYTIQ.Y_QUALIFPIECE) AND A1.Y_AXE';
        SQLA := SQLA + ' = ANALYTIQ.Y_AXE AND A1.Y_NUMVENTIL = ANALYTIQ.Y_NUMVENTIL AND A1.Y_QUALIFPIECE = ANALYTIQ.Y_QUALIFPIECE) WHERE EXISTS (SELECT E_EXERCICE FROM ECRITURE  WHERE E_JOURNAL = Y_JOURNAL';
        SQLA := SQLA + ' AND E_EXERCICE = Y_EXERCICE AND E_DATECOMPTABLE = Y_DATECOMPTABLE AND E_NUMEROPIECE = Y_NUMEROPIECE AND E_NUMLIGNE = Y_NUMLIGNE AND E_QUALIFPIECE = Y_QUALIFPIECE AND E_MODESAISIE="-")';
        try
            ExecuteSQL(SQLA);
        except
            PGIBox(erreur);
            RollBack;
            SourisNormale;
            EnableControls(Self,true);
            exit;
        end;
    end;

    // date comptable
    if (CBDate.Checked) then
    begin
        // ECRITURE
        SQLE := 'UPDATE ECRITURE SET';
        SQLE := SQLE + ' E_DATECOMPTABLE = (SELECT E1.E_DATECOMPTABLE FROM ECRITURE E1 WHERE';
        SQLE := SQLE + ' E1.E_JOURNAL = ECRITURE.E_JOURNAL AND E1.E_EXERCICE = ECRITURE.E_EXERCICE AND E1.E_NUMEROPIECE = ECRITURE.E_NUMEROPIECE AND E1.E_NUMLIGNE = (SELECT';
        SQLE := SQLE + ' MAX(E2.E_NUMLIGNE) FROM ECRITURE E2 WHERE E2.E_JOURNAL = ECRITURE.E_JOURNAL AND E2.E_EXERCICE = ECRITURE.E_EXERCICE AND E2.E_NUMEROPIECE = ECRITURE.E_NUMEROPIECE';
        SQLE := SQLE + ' AND E2.E_QUALIFPIECE = ECRITURE.E_QUALIFPIECE) AND E1.E_NUMECHE IN (0,1) AND E1.E_QUALIFPIECE = ECRITURE.E_QUALIFPIECE) WHERE E_MODESAISIE="-"';
        try
            ExecuteSQL(SQLE);
        except
            PGIBox(erreur);
            RollBack;
            SourisNormale;
            EnableControls(Self,true);
            exit;
        end;
        // ANALYTIQUE
        SQLA := 'UPDATE ANALYTIQ SET';
        SQLA := SQLA + ' Y_DATECOMPTABLE = (SELECT A1.Y_DATECOMPTABLE FROM ANALYTIQ A1 WHERE';
        SQLA := SQLA + ' A1.Y_JOURNAL = ANALYTIQ.Y_JOURNAL AND A1.Y_EXERCICE = ANALYTIQ.Y_EXERCICE AND A1.Y_NUMEROPIECE = ANALYTIQ.Y_NUMEROPIECE AND A1.Y_NUMLIGNE = (SELECT';
        SQLA := SQLA + ' MAX(A2.Y_NUMLIGNE) FROM ANALYTIQ A2 WHERE A2.Y_JOURNAL = ANALYTIQ.Y_JOURNAL AND A2.Y_EXERCICE = ANALYTIQ.Y_EXERCICE AND A2.Y_NUMEROPIECE = ANALYTIQ.Y_NUMEROPIECE';
        SQLA := SQLA + ' AND A2.Y_AXE = ANALYTIQ.Y_AXE AND A2.Y_NUMVENTIL = ANALYTIQ.Y_NUMVENTIL AND A2.Y_QUALIFPIECE =  ANALYTIQ.Y_QUALIFPIECE) AND A1.Y_AXE = ANALYTIQ.Y_AXE AND ';
        SQLA := SQLA + ' A1.Y_NUMVENTIL = ANALYTIQ.Y_NUMVENTIL AND A1.Y_QUALIFPIECE = ANALYTIQ.Y_QUALIFPIECE) WHERE EXISTS (SELECT E_EXERCICE FROM ECRITURE  WHERE E_JOURNAL = Y_JOURNAL';
        SQLA := SQLA + ' AND E_EXERCICE = Y_EXERCICE AND E_NUMEROPIECE = Y_NUMEROPIECE AND E_NUMLIGNE = Y_NUMLIGNE AND E_QUALIFPIECE = Y_QUALIFPIECE AND E_MODESAISIE="-")';
        try
            ExecuteSQL(SQLA);
        except
            PGIBox(erreur);
            RollBack;
            SourisNormale;
            EnableControls(Self,true);
            exit;
        end;
    end;

    if CPiece.Checked and not ReparePiece then
     begin
      RollBack ;
      exit;
     end ;

    CommitTrans;

    SourisNormale;
    EnableControls(Self,true);

    Close;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 15/06/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFRepEntete.BAideClick(Sender: TObject);
begin
    CallHelpTopic(Self) ;
end;

//================================================================================
// fin
//================================================================================
end.
