{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 21/10/2003
Modifié le ... :   /  /
Description .. : Edition déclaration 2483 formation
Mots clefs ... : PAIE;FORMATION
*****************************************************************
PT1  | 16/11/2004 | V_60  | JL | Correction violation d'accès
PT2  | 11/01/2004 | V_60  | JL | Correction requete (PSS_COUTPEDAG absent du GROUP BY)
PT3  | 15/03/2005 | V_60  | JL | FQ 12092 Correction requete oracle
---  | 20/03/2006 |       | JL | modification clé annuaire ----
PT4  | 05/10/2006 | V_75  | JL | Corrections calcul cout interne
PT5  | 01/06/2007 | V_72  | FL | FQ xxxxx Ajout de bouton d'aperçu pour certains calculs
PT6  | 11/02/2008 | V_804 | FL | Mise à jour pour la déclaration de 2007
}
unit UTofPgEditDeclaration2483;

interface
uses StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
     eQRS1,UtilEAgl,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}QRS1,EdtREtat,
{$ENDIF}
    forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,UTob,HSysMenu,HTB97,ParamDat,EntPaie,PARAMSOC,Menus ;

Type
	TOF_PGDECLARAFORMATION = Class (TOF)
	private
		millesime : String;
		DateDeb,DateFin : TDateTime; //PT5
		procedure EditDeclaration(Sender: TObject);
		procedure CalculerDecl(Sender: TObject);
		procedure ValiderDecl(Sender : TObject);
		procedure ChargerDecl;
		Procedure EditDetailsA(Sender : TObject); //PT5
		Procedure EditDetailsB(Sender : TObject); //PT5
		Procedure EditDetailsD(Sender : TObject); //PT5
		Procedure CalcTotal (Sender : TObject);   //PT6
		Procedure CalcMontant (Sender : TObject); //PT6
        Procedure CalcExcedents (Sender : TObject); //PT6
	public
		procedure OnArgument(Arguments : String ) ; override ;
	END ;

implementation

Uses PGOutilsFormation, PGOutils2, StrUtils, DateUtils, HPDFViewer;

procedure TOF_PGDECLARAFORMATION.CalculerDecl(Sender : TObject);
var TobEtat,T,TobFormations,TobSalaries,TS,TobEtablissements,TobLesDepenses,TobInvest : Tob;
    Q : TQuery;
    i : Integer;
    Pages : TPageControl;
    OPCAInterne,StPages : String;
    TabAGM,TabOnQ,TabOQ,TabEmp,TabIC : array  [1..12] of Double;
    PrisEnCharge : Boolean;
    NbOnqH,NbOnqF,NbOQH,NbOQF,NbEmpH,NbEmpF,NbAgmH,NbAgmF,NbICH,NbICF : Double;
    Etab1,Etab2,Etab3 : Integer;
    DepensesInt,DepensesAnn,DepensesPlu,DepensesBdC,Total : Double;
    CoutOrgan,SalaireSal,SalaireA,FraisSal,CoutMob,CoutFonct,CoutPedag : Double;
    DateAge1,DateAge2,DateDebN1 : TDateTime;
    CatDucs,CatDads : String;
    MasseSal,MasseSalCDD,VerseCIF,VersePro,VersePlan,VerseALT : Double;
    Frais,CoutUnit,NbSta : Double;
    //totH,Totf,TotDIF,TotDIFH,TotPLF : Double; //PT6
    NbH,NbF,Heures,MtCumul : Double;
    CumulDIF : String;
    Alloc,DepensesExt : Double;
    WhereMulti,WhereMultiS : String; //PT6
    DepExtFor, DepExtVAE, DepExtBIL, ExcA,ExcB,ExcC,ExcD,ExcE : Double; //PT6
    MillesimeOld : String; //PT6
begin
        CumulDIF := GetParamSOc('SO_PGCUMULDIFACQUIS');
        MasseSal := 0;
        //MasseSalCDD := 0; //PT6
        DateDeb := IDate1900;
        DateFin := IDate1900;
        DateDebN1 := IDate1900; //PT6
        Pages := TPageControl(GetControl('PAGES'));
        StPages := '';
        OPCAInterne := VH_Paie.PGForOPCAInterne;
        
        //PT6 - Début
        If PGBundleInscFormation then
        Begin
        	WhereMulti := ' AND ((PFO_PREDEFINI="DOS" AND PFO_NODOSSIER="'+PGRendNoDossier+'") OR (PFO_PREDEFINI="STD"))';
        	WhereMultiS := ' AND ((PSS_PREDEFINI="DOS" AND PSS_NODOSSIER="'+PGRendNoDossier+'") OR (PSS_PREDEFINI="STD"))';
        End
        Else
        Begin
			WhereMulti := '';
			WhereMultiS := '';
		End;
        //PT6 - Fin
        
        Q := OpenSQL('SELECT PFE_MASSESAL,PFE_MASSESALCDD,PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="'+Millesime+'"',True);
        If Not Q.Eof Then
        begin
                DateDeb := Q.FindField('PFE_DATEDEBUT').AsDateTime;
                DateFin := Q.FindField('PFE_DATEFIN').AsDateTime;
                MasseSal := Q.FindField('PFE_MASSESAL').AsFloat;
                //MasseSalCDD := Q.FindField('PFE_MASSESALCDD').AsFloat; //PT6
				DateDebN1 := IncYear(DateDeb,-1);
        end;
        Ferme(Q);
        
        // Calcul des efectifs par catégories et sexes
        //TotH := 0; //PT6
        //TotF := 0;
        Q := OpenSQL('SELECT SUM(PSA_UNITEPRISEFF) NBSAL FROM SALARIES '+
        'WHERE PSA_PRISEFFECTIF="X" AND PSA_SEXE="M" AND PSA_CATDADS="004" AND PSA_DATESORTIE="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>"'+UsDateTime(DateFin)+'"',True);
        If Not Q.Eof then NbOnqH := Q.FindField('NBSAL').AsFloat
        else NbOnQH := 0;
        SetControlText('B2A',FloatToStr(NbOnQH));
        //TotH := TotH + NbOnQH; //PT6
        Ferme(Q);
        
        Q := OpenSQL('SELECT SUM(PSA_UNITEPRISEFF) NBSAL FROM SALARIES '+
        'WHERE PSA_PRISEFFECTIF="X" AND PSA_SEXE="F" AND PSA_CATDADS="004" AND PSA_DATESORTIE="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>"'+UsDateTime(DateFin)+'"',True);
        If Not Q.Eof then NbOnqF := Q.FindField('NBSAL').AsFloat
        else NbOnqF := 0;
        Ferme(Q);
        SetControlText('B2B',FloatToStr(NbOnqF));
        
        //TotF := TotF + NbOnqF; //PT6
        Q := OpenSQL('SELECT SUM(PSA_UNITEPRISEFF) NBSAL FROM SALARIES '+
        'WHERE PSA_PRISEFFECTIF="X" AND PSA_SEXE="M" AND PSA_CATDADS="005" AND (PSA_DADSPROF="" OR PSA_DADSPROF="02") AND PSA_DATESORTIE="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>"'+UsDateTime(DateFin)+'"',True);
        If Not Q.Eof then NbEmpH := Q.FindField('NBSAL').AsFloat
        else NbEmpH := 0;
        Ferme(Q);
        SetControlText('B3A',FloatToStr(NbEmpH));
        
        //TotH := TotH + NbEmpH; //PT6
        Q := OpenSQL('SELECT SUM(PSA_UNITEPRISEFF) NBSAL FROM SALARIES '+
        'WHERE PSA_PRISEFFECTIF="X" AND PSA_SEXE="F" AND PSA_CATDADS="005" AND (PSA_DADSPROF="" OR PSA_DADSPROF="02") AND PSA_DATESORTIE="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>"'+UsDateTime(DateFin)+'"',True);
        If Not Q.Eof then NbEmpF := Q.FindField('NBSAL').AsFloat
        else NbEmpF := 0;
        Ferme(Q);
        SetControlText('B3B',FloatToStr(NbEmpF));
        
        //TotF := TotF + NbEmpF; //PT6
        Q := OpenSQL('SELECT SUM(PSA_UNITEPRISEFF) NBSAL FROM SALARIES '+
        'WHERE PSA_PRISEFFECTIF="X" AND PSA_SEXE="M" AND PSA_CATDADS="005" AND (PSA_DADSPROF<>"" AND PSA_DADSPROF<>"02") AND PSA_DATESORTIE="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>"'+UsDateTime(DateFin)+'"',True);
        If Not Q.Eof then NbAgmH := Q.FindField('NBSAL').AsFloat
        else NbAgmH := 0;
        Ferme(Q);
        SetControlText('B4A',FloatToStr(NbAgmH));
        
        //TotH := TotH + NbAgmH; //PT6
        Q := OpenSQL('SELECT SUM(PSA_UNITEPRISEFF) NBSAL FROM SALARIES '+
        'WHERE PSA_PRISEFFECTIF="X" AND PSA_SEXE="F" AND PSA_CATDADS="005" AND (PSA_DADSPROF<>"" AND PSA_DADSPROF<>"02") AND PSA_DATESORTIE="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>"'+UsDateTime(DateFin)+'"',True);
        If Not Q.Eof then NbAgmF := Q.FindField('NBSAL').AsFloat
        else NbAgmF := 0;
        Ferme(Q);
        SetControlText('B4B',FloatToStr(NbAgmF));
        
        //TotF := TotF + NbAgmF; //PT6
        Q := OpenSQL('SELECT SUM(PSA_UNITEPRISEFF) NBSAL FROM SALARIES '+
        'WHERE PSA_PRISEFFECTIF="X" AND PSA_SEXE="M" AND (PSA_CATDADS="001" OR PSA_CATDADS="002") AND PSA_DATESORTIE="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>"'+UsDateTime(DateFin)+'"',True);
        If Not Q.Eof then NbICH := Q.FindField('NBSAL').AsFloat
        else NbICH := 0;
        Ferme(Q);
        SetControlText('B5A',FloatToStr(NbICH));
        
        //TotH := TotH + NbICH; //PT6
        Q := OpenSQL('SELECT SUM(PSA_UNITEPRISEFF) NBSAL FROM SALARIES '+
        'WHERE PSA_PRISEFFECTIF="X" AND PSA_SEXE="F" AND (PSA_CATDADS="001" OR PSA_CATDADS="002") AND PSA_DATESORTIE="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>"'+UsDateTime(DateFin)+'"',True);
        If Not Q.Eof then NbICF := Q.FindField('NBSAL').AsFloat
        else NbICF := 0;
        Ferme(Q);
        SetControlText('B5B',FloatToStr(NbICF));
        
        //TotF := TotF + NbICF; //PT6
        //SetControlText('B6A',FloatToStr(TotH));
        //SetControlText('B6B',FloatToStr(TotF));
        //PLAN DE FORMATION
        //TotPLF := 0; //PT6
        //TotH := 0;
        //TotF := 0;
        Q := OpenSQL('SELECT COUNT(PFO_SALARIE) NBSAL,SUM(PFO_NBREHEURE) NBHEURES  FROM FORMATIONS '+
        'LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE '+
        'LEFT JOIN SALARIES ON PFO_SALARIE=PSA_SALARIE '+
        'WHERE PSA_SEXE="M" AND PSA_CATDADS="004" '+
        'AND PFO_TYPEPLANPREV="PLF" AND PFO_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PFO_DATEDEBUT<="'+UsDateTime(DateFin)+'" AND PSS_INCLUSDECL="X" AND PFO_EFFECTUE="X"'+WhereMulti,True); //PT6
        If Not Q.Eof then
        begin
             NbH := Q.FindField('NBSAL').AsFloat;
             Heures :=Q.FindField('NBHEURES').AsFloat;
        end
        else
        begin
             NbH := 0;
             Heures := 0;
        end;
        Ferme(Q);
        
        Q := OpenSQL('SELECT COUNT(PFO_SALARIE) NBSAL,SUM(PFO_NBREHEURE) NBHEURES  FROM FORMATIONS '+
        'LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE '+
        'LEFT JOIN SALARIES ON PFO_SALARIE=PSA_SALARIE '+
        'WHERE PSA_SEXE="F" AND PSA_CATDADS="004" '+
        'AND PFO_TYPEPLANPREV="PLF" AND (PFO_DATEFIN>="'+UsDateTime(DateDeb)+'" AND PFO_DATEFIN<="'+UsDateTime(DateFin)+'") AND PFO_DATEDEBUT>="'+UsDateTime(DateDebN1)+ //PT6
        '" AND PSS_INCLUSDECL="X" AND PFO_EFFECTUE="X"'+WhereMulti,True); //PT6
        If Not Q.Eof then
        begin
             NbF := Q.FindField('NBSAL').AsFloat;
             Heures :=Heures + Q.FindField('NBHEURES').AsFloat;
        end
        else NbF := 0;
        Ferme(Q);
        
        SetControlText('B2C',FloatToStr(NbH));
        SetControlText('B2D',FloatToStr(NbF));
        SetControlText('B2E',FloatToStr(Heures));
        //TotPLF := totPlf + Heures; //PT6
        //TotH := TotH + NbH;
        //TotF := TotF+NbF;
        
        Q := OpenSQL('SELECT COUNT(PFO_SALARIE) NBSAL,SUM(PFO_NBREHEURE) NBHEURES  FROM FORMATIONS '+
        'LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE '+
        'LEFT JOIN SALARIES ON PFO_SALARIE=PSA_SALARIE '+
        'WHERE PSA_SEXE="M" AND PSA_CATDADS="005" AND (PSA_DADSPROF="" OR PSA_DADSPROF="02") '+
        //'AND PFO_TYPEPLANPREV="PLF" AND PFO_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PFO_DATEDEBUT<="'+UsDateTime(DateFin)+
        'AND PFO_TYPEPLANPREV="PLF" AND (PFO_DATEFIN>="'+UsDateTime(DateDeb)+'" AND PFO_DATEFIN<="'+UsDateTime(DateFin)+'") AND PFO_DATEDEBUT>="'+UsDateTime(DateDebN1)+ //PT6
        '" AND PSS_INCLUSDECL="X" AND PFO_EFFECTUE="X"'+WhereMulti,True); //PT6
        If Not Q.Eof then
        begin
             NbH := Q.FindField('NBSAL').AsFloat;
             Heures :=Q.FindField('NBHEURES').AsFloat;
        end
        else
        begin
             NbH := 0;
             Heures := 0;
        end;
        Ferme(Q);
        
        Q := OpenSQL('SELECT COUNT(PFO_SALARIE) NBSAL,SUM(PFO_NBREHEURE) NBHEURES  FROM FORMATIONS '+
        'LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE '+
        'LEFT JOIN SALARIES ON PFO_SALARIE=PSA_SALARIE '+
        'WHERE PSA_SEXE="F" AND PSA_CATDADS="005" AND (PSA_DADSPROF="" OR PSA_DADSPROF="02") '+
        //'AND PFO_TYPEPLANPREV="PLF" AND PFO_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PFO_DATEDEBUT<="'+UsDateTime(DateFin)+
        'AND PFO_TYPEPLANPREV="PLF" AND (PFO_DATEFIN>="'+UsDateTime(DateDeb)+'" AND PFO_DATEFIN<="'+UsDateTime(DateFin)+'") AND PFO_DATEDEBUT>="'+UsDateTime(DateDebN1)+ //PT6
        '" AND PSS_INCLUSDECL="X" AND PFO_EFFECTUE="X"'+WhereMulti,True); //PT6
        If Not Q.Eof then
        begin
             NbF := Q.FindField('NBSAL').AsFloat;
             Heures :=Heures + Q.FindField('NBHEURES').AsFloat;
        end
        else NbF := 0;
        Ferme(Q);
        
        SetControlText('B3C',FloatToStr(NbH));
        SetControlText('B3D',FloatToStr(NbF));
        SetControlText('B3E',FloatToStr(Heures));
        //TotPLF := totPlf + Heures; //PT6
        //TotH := TotH + NbH;
        //TotF := TotF+NbF;
        
        Q := OpenSQL('SELECT COUNT(PFO_SALARIE) NBSAL,SUM(PFO_NBREHEURE) NBHEURES FROM FORMATIONS '+
        'LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE '+
        'LEFT JOIN SALARIES ON PFO_SALARIE=PSA_SALARIE '+
        'WHERE PSA_SEXE="M" AND PSA_CATDADS="005" AND (PSA_DADSPROF<>"" AND PSA_DADSPROF<>"02") '+
        //'AND PFO_TYPEPLANPREV="PLF" AND PFO_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PFO_DATEDEBUT<="'+UsDateTime(DateFin)+
        'AND PFO_TYPEPLANPREV="PLF" AND (PFO_DATEFIN>="'+UsDateTime(DateDeb)+'" AND PFO_DATEFIN<="'+UsDateTime(DateFin)+'") AND PFO_DATEDEBUT>="'+UsDateTime(DateDebN1)+ //PT6
        '" AND PSS_INCLUSDECL="X" AND PFO_EFFECTUE="X"'+WhereMulti,True); //PT6
        If Not Q.Eof then
        begin
             NbH := Q.FindField('NBSAL').AsFloat;
             Heures :=Q.FindField('NBHEURES').AsFloat;
        end
        else
        begin
             NbH := 0;
             Heures := 0;
        end;
        Ferme(Q);
        
        Q := OpenSQL('SELECT COUNT(PFO_SALARIE) NBSAL,SUM(PFO_NBREHEURE) NBHEURES FROM FORMATIONS '+
        'LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE '+
        'LEFT JOIN SALARIES ON PFO_SALARIE=PSA_SALARIE '+
        'WHERE PSA_SEXE="F" AND PSA_CATDADS="005" AND (PSA_DADSPROF<>"" AND PSA_DADSPROF<>"02") '+
        //'AND PFO_TYPEPLANPREV="PLF" AND PFO_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PFO_DATEDEBUT<="'+UsDateTime(DateFin)+
        'AND PFO_TYPEPLANPREV="PLF" AND (PFO_DATEFIN>="'+UsDateTime(DateDeb)+'" AND PFO_DATEFIN<="'+UsDateTime(DateFin)+'") AND PFO_DATEDEBUT>="'+UsDateTime(DateDebN1)+ //PT6
        '" AND PSS_INCLUSDECL="X" AND PFO_EFFECTUE="X"'+WhereMulti,True); //PT6
        If Not Q.Eof then
        begin
             NbF := Q.FindField('NBSAL').AsFloat;
             Heures :=Heures + Q.FindField('NBHEURES').AsFloat;
        end
        else NbF := 0;
        Ferme(Q);
        
        SetControlText('B4C',FloatToStr(NbH));
        SetControlText('B4D',FloatToStr(NbF));
        SetControlText('B4E',FloatToStr(Heures));
        //TotPLF := totPlf + Heures; //PT6
        //TotH := TotH + NbH;
        //TotF := TotF+NbF;
        
        Q := OpenSQL('SELECT COUNT(PFO_SALARIE) NBSAL,SUM(PFO_NBREHEURE) NBHEURES  FROM FORMATIONS '+
        'LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE '+
        'LEFT JOIN SALARIES ON PFO_SALARIE=PSA_SALARIE '+
        'WHERE PSA_SEXE="M" AND (PSA_CATDADS="001" OR PSA_CATDADS="002") '+
        //'AND PFO_TYPEPLANPREV="PLF" AND PFO_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PFO_DATEDEBUT<="'+UsDateTime(DateFin)+
        'AND PFO_TYPEPLANPREV="PLF" AND (PFO_DATEFIN>="'+UsDateTime(DateDeb)+'" AND PFO_DATEFIN<="'+UsDateTime(DateFin)+'") AND PFO_DATEDEBUT>="'+UsDateTime(DateDebN1)+ //PT6
        '" AND PSS_INCLUSDECL="X" AND PFO_EFFECTUE="X"'+WhereMulti,True); //PT6
        If Not Q.Eof then
        begin
             NbH := Q.FindField('NBSAL').AsFloat;
             Heures :=Q.FindField('NBHEURES').AsFloat;
        end
        else
        begin
             NbH := 0;
             Heures := 0;
        end;
        Ferme(Q);
        
        Q := OpenSQL('SELECT COUNT(PFO_SALARIE) NBSAL,SUM(PFO_NBREHEURE) NBHEURES  FROM FORMATIONS '+
        'LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE '+
        'LEFT JOIN SALARIES ON PFO_SALARIE=PSA_SALARIE '+
        'WHERE PSA_SEXE="F" AND (PSA_CATDADS="001" OR PSA_CATDADS="002") '+
        //'AND PFO_TYPEPLANPREV="PLF" AND PFO_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PFO_DATEDEBUT<="'+UsDateTime(DateFin)+
        'AND PFO_TYPEPLANPREV="PLF" AND (PFO_DATEFIN>="'+UsDateTime(DateDeb)+'" AND PFO_DATEFIN<="'+UsDateTime(DateFin)+'") AND PFO_DATEDEBUT>="'+UsDateTime(DateDebN1)+ //PT6
        '" AND PSS_INCLUSDECL="X" AND PFO_EFFECTUE="X"'+WhereMulti,True); //PT6
        If Not Q.Eof then
        begin
             NbF := Q.FindField('NBSAL').AsFloat;
             Heures :=Heures + Q.FindField('NBHEURES').AsFloat;
        end
        else NbF := 0;
        Ferme(Q);
        
        SetControlText('B5C',FloatToStr(NbH));
        SetControlText('B5D',FloatToStr(NbF));
        SetControlText('B5E',FloatToStr(Heures));
        //TotPLF := totPlf + Heures; //PT6
        //TotH := TotH + NbH;
        //TotF := TotF+NbF;

        //SetControlText('B6C',FloatToStr(TotH)); //PT6
        //SetControlText('B6D',FloatToStr(TotF));
        //SetControlText('B6E',FloatToStr(TotPLF));
        // DIF
        //TotDIF := 0; //PT6
        //TotDIFH := 0;
        Q := OpenSQL('SELECT COUNT(PFO_SALARIE) NBSAL,SUM(PFO_NBREHEURE) NBHEURES  FROM FORMATIONS '+
        'LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE '+
        'LEFT JOIN SALARIES ON PFO_SALARIE=PSA_SALARIE '+
        'WHERE PSA_CATDADS="004" '+
        'AND PFO_TYPEPLANPREV="DIF" AND PFO_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PFO_DATEDEBUT<="'+UsDateTime(DateFin)+'" AND PSS_INCLUSDECL="X" AND PFO_EFFECTUE="X"'+WhereMulti,True); //PT6
        If Not Q.Eof then NbOnqH := Q.FindField('NBSAL').AsFloat
        else NbOnQH := 0;
        SetControlText('B2F',FloatToStr(NbOnQH));
        //TotDIF := TotDIF + NbOnQH; //PT6
        If Not Q.Eof then NbOnqH := Q.FindField('NBHEURES').AsFloat
        else NbOnQH := 0;
        SetControlText('B2G',FloatToStr(NbOnQH));
        //TotDIFH := TotDIFH + NbOnQH; //PT6
        Ferme(Q);
        
        Q := OpenSQL('SELECT COUNT(PFO_SALARIE) NBSAL,SUM(PFO_NBREHEURE) NBHEURES  FROM FORMATIONS '+
        'LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE '+
        'LEFT JOIN SALARIES ON PFO_SALARIE=PSA_SALARIE '+
        'WHERE PSA_CATDADS="005" AND (PSA_DADSPROF="" OR PSA_DADSPROF="02") '+
        'AND PFO_TYPEPLANPREV="DIF" AND PFO_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PFO_DATEDEBUT<="'+UsDateTime(DateFin)+'" AND PSS_INCLUSDECL="X" AND PFO_EFFECTUE="X"'+WhereMulti,True); //PT6
        If Not Q.Eof then NbEmpH := Q.FindField('NBSAL').AsFloat
        else NbEmpH := 0;
        SetControlText('B3F',FloatToStr(NbEmpH));
        //TotDIF := TotDIF + NbEmpH; //PT6
        If Not Q.Eof then NbOnqH := Q.FindField('NBHEURES').AsFloat
        else NbOnQH := 0;
        Ferme(Q);
        
        SetControlText('B3G',FloatToStr(NbOnQH));
        //TotDIFH := TotDIFH + NbOnQH; //PT6
        Q := OpenSQL('SELECT COUNT(PFO_SALARIE) NBSAL,SUM(PFO_NBREHEURE) NBHEURES FROM FORMATIONS '+
        'LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE '+
        'LEFT JOIN SALARIES ON PFO_SALARIE=PSA_SALARIE '+
        'WHERE PSA_CATDADS="005" AND (PSA_DADSPROF<>"" AND PSA_DADSPROF<>"02") '+
        'AND PFO_TYPEPLANPREV="DIF" AND PFO_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PFO_DATEDEBUT<="'+UsDateTime(DateFin)+'" AND PSS_INCLUSDECL="X" AND PFO_EFFECTUE="X"'+WhereMulti,True); //PT6
        If Not Q.Eof then NbAgmH := Q.FindField('NBSAL').AsFloat
        else NbAgmH := 0;
        SetControlText('B4F',FloatToStr(NbAgmH));
        //TotDIF := TotDIF + NbAgmH; //PT6
        If Not Q.Eof then NbOnqH := Q.FindField('NBHEURES').AsFloat
        else NbOnQH := 0;
        Ferme(Q);
        
        SetControlText('B4G',FloatToStr(NbOnQH));
        //TotDIFH := TotDIFH + NbOnQH; //PT6
        Q := OpenSQL('SELECT COUNT(PFO_SALARIE) NBSAL,SUM(PFO_NBREHEURE) NBHEURES  FROM FORMATIONS '+
        'LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE '+
        'LEFT JOIN SALARIES ON PFO_SALARIE=PSA_SALARIE '+
        'WHERE (PSA_CATDADS="001" OR PSA_CATDADS="002") '+
        'AND PFO_TYPEPLANPREV="DIF" AND PFO_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PFO_DATEDEBUT<="'+UsDateTime(DateFin)+'" AND PSS_INCLUSDECL="X" AND PFO_EFFECTUE="X"'+WhereMulti,True); //PT6
        If Not Q.Eof then NbICH := Q.FindField('NBSAL').AsFloat
        else NbICH := 0;
        SetControlText('B5F',FloatToStr(NbICH));
        //TotDIF := TotDIF + NbICH; //PT6
        If Not Q.Eof then NbOnqH := Q.FindField('NBHEURES').AsFloat
        else NbOnQH := 0;
        Ferme(Q);
        
        SetControlText('B5G',FloatToStr(NbOnQH));
        //TotDIFH := TotDIFH + NbOnQH; //PT6
        //SetControlText('B6F',FloatToStr(TotDIF)); //PT6
        //SetControlText('B6G',FloatToStr(TotDIFH));
        //TotDIF := 0;
        
        Q := OpenSQL('SELECT SUM(PHC_MONTANT) ACQUIS FROM HISTOCUMSAL'+
        ' LEFT JOIN SALARIES ON PHC_SALARIE=PSA_SALARIE '+
        'WHERE PSA_CATDADS="004" '+
        'AND PHC_CUMULPAIE="'+CumulDIF+'"',True);
        If Not Q.Eof then MtCumul := Q.FindField('ACQUIS').AsFloat
        else MtCumul := 0;
        SetControlText('B2H',FloatToStr(MtCumul));
        //TotDIF := TotDIF + MtCumul; //PT6
        Ferme(Q); //PT6
        
        Q := OpenSQL('SELECT SUM(PHC_MONTANT) ACQUIS FROM HISTOCUMSAL'+
        ' LEFT JOIN SALARIES ON PHC_SALARIE=PSA_SALARIE '+
        'WHERE PSA_CATDADS="005" AND (PSA_DADSPROF="" OR PSA_DADSPROF="02") '+
        'AND PHC_CUMULPAIE="'+CumulDIF+'"',True);
        If Not Q.Eof then MtCumul := Q.FindField('ACQUIS').AsFloat
        else MtCumul := 0;
        SetControlText('B3H',FloatToStr(MtCumul));
        //TotDIF := TotDIF + MtCumul; //PT6
        Ferme(Q); //PT6
        
        Q := OpenSQL('SELECT SUM(PHC_MONTANT) ACQUIS FROM HISTOCUMSAL'+
        ' LEFT JOIN SALARIES ON PHC_SALARIE=PSA_SALARIE '+
        'WHERE PSA_CATDADS="005" AND (PSA_DADSPROF<>"" AND PSA_DADSPROF<>"02") '+
        'AND PHC_CUMULPAIE="'+CumulDIF+'"',True);
        If Not Q.Eof then MtCumul := Q.FindField('ACQUIS').AsFloat
        else MtCumul := 0;
        SetControlText('B4H',FloatToStr(MtCumul));
        //TotDIF := TotDIF + MtCumul; //PT6
        Ferme(Q); //PT6
        
        Q := OpenSQL('SELECT SUM(PHC_MONTANT) ACQUIS FROM HISTOCUMSAL'+
        ' LEFT JOIN SALARIES ON PHC_SALARIE=PSA_SALARIE '+
        'WHERE (PSA_CATDADS="001" OR PSA_CATDADS="002") '+
        'AND PHC_CUMULPAIE="'+CumulDIF+'"',True);
        If Not Q.Eof then MtCumul := Q.FindField('ACQUIS').AsFloat
        else MtCumul := 0;
        SetControlText('B5H',FloatToStr(MtCumul));
        //TotDIF := TotDIF + MtCumul; //PT6
        Ferme(Q); //PT6
        
        //SetControlText('B6H',FloatToStr(TotDif)); //PT6
        //période de professionalisation
        Q := OpenSQL('SELECT COUNT(PFO_SALARIE) NBSAL,SUM(PFO_NBREHEURE) NBHEURES FROM FORMATIONS '+
        'LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE '+
        'WHERE  PFO_TYPEPLANPREV="PRO" AND PFO_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PFO_DATEDEBUT<="'+UsDateTime(DateFin)+'" AND PSS_INCLUSDECL="X" AND PFO_EFFECTUE="X"'+WhereMulti,True); //PT6
        If Not Q.Eof then
        begin
                SetControlText('B7',IntToStr(Q.FindField('NBSAL').AsInteger));
                SetControlText('B8',FloatToStr(Q.FindField('NBHEURES').AsFloat));
        end
        else
        begin
                SetControlText('B7','0');
                SetControlText('B8','0');
        end;
        Ferme(Q);
        
        // Allocations de formation
        Q := OpenSQL('SELECT COUNT(PFO_SALARIE) NBSAL,SUM(PFO_NBREHEURE) NBHEURES FROM FORMATIONS '+
        'LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE '+
        'WHERE  PFO_ALLOCFORM>0 AND PFO_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PFO_DATEDEBUT<="'+UsDateTime(DateFin)+'" AND PSS_INCLUSDECL="X" AND PFO_EFFECTUE="X"'+WhereMulti,True); //PT6
        If Not Q.Eof then
        begin
                SetControlText('B9',IntToStr(Q.FindField('NBSAL').AsInteger));
                SetControlText('B10',FloatToStr(Q.FindField('NBHEURES').AsFloat));
        end
        else
        begin
                SetControlText('B9','0');
                SetControlText('B10','0');
        end;
        Ferme(Q);
        
        // Nombre de bilans de compétence //PT6
        Q := OpenSQL('SELECT COUNT(*) NBBILANS FROM FORMATIONS '+
        'WHERE PFO_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PFO_DATEDEBUT<="'+UsDateTime(DateFin)+
        '" AND PFO_INCLUSDECL="X" AND PFO_EFFECTUE="X" AND PFO_ACTIONFORM="BIL"'+WhereMulti,True); 
        If Not Q.Eof then
                SetControlText('B11',IntToStr(Q.FindField('NBBILANS').AsInteger))
        else
                SetControlText('B11','0');
        Ferme(Q);
        
        // Nombre de validations des acquis //PT6
        Q := OpenSQL('SELECT COUNT(*) NBVAE FROM FORMATIONS '+
        'WHERE PFO_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PFO_DATEDEBUT<="'+UsDateTime(DateFin)+
        '" AND PFO_INCLUSDECL="X" AND PFO_EFFECTUE="X" AND PFO_ACTIONFORM="VAE"'+WhereMulti,True); 
        If Not Q.Eof then
                SetControlText('B12',IntToStr(Q.FindField('NBVAE').AsInteger))
        else
                SetControlText('B12','0');
        Ferme(Q);
        
        //MASSE SALARIALE
{        Q := OpenSQL('SELECT SUM(PPU_CBRUT) MASSESAL FROM PAIEENCOURS LEFT JOIN CONTRATTRAVAIL '+
        'ON PPU_SALARIE=PCI_SALARIE AND PPU_DATEDEBUT>=PCI_DEBUTCONTRAT AND PPU_DATEDEBUT<=PCI_FINCONTRAT '+
        'WHERE PCI_TYPECONTRAT<>"CCD" AND PPU_DATEDEBUT>="'+UsDateTime(DateDeb)+'" '+
        'AND PPU_DATEFIN<="'+UsDateTime(Datefin)+'" ',True);
        If Not Q.eof then MasseSal := Q.FindField('MASSESAL').AsFloat
        else MasseSal := 0;
        Ferme(Q);}
        SetControlText('MONTANT1',IntToStr((Round(MasseSal)))); //PT6
        
        Q := OpenSQL('SELECT SUM(PIF_POURCENTAGE) MONTANT FROM INVESTFORMATION '+
        'WHERE PIF_ORGCOLLECTGU<>"'+OPCAInterne+'" AND PIF_MILLESIME="'+millesime+'" AND PIF_INVESTFORM="PLF"',True);
        If Not Q.Eof then VersePro := Q.FindField('MONTANT').AsFloat
        else VersePro := 0;
        Ferme(Q);
        SetControlText('MONTANT2',FloatToStr(VersePro)); //PT6

        //CIF
        Q := OpenSQL('SELECT SUM(PIF_POURCENTAGE) MONTANT FROM INVESTFORMATION '+
        'WHERE PIF_ORGCOLLECTGU<>"'+OPCAInterne+'" AND PIF_MILLESIME="'+millesime+'" AND PIF_INVESTFORM="CIF"',True);
        If Not Q.Eof then VerseCIF := Q.FindField('MONTANT').AsFloat
        else VerseCIF := 0;
        Ferme(Q);
        SetControlText('MONTANT3',FloatToStr(VerseCIF)); //PT6
        
        //SetControlText('MONTANT4',FloatToStr(StrToFloat(GetControlText('MONTANT1')) * VerseCIF/100)); //PT6
        Q := OpenSQL('SELECT SUM(PIF_MTVERSEMENT) MONTANT FROM INVESTFORMATION '+
        'WHERE PIF_ORGCOLLECTGU<>"'+OPCAInterne+'" AND PIF_MILLESIME="'+millesime+'" AND PIF_INVESTFORM="CIF"',True);
        If Not Q.Eof then VerseCIF := Q.FindField('MONTANT').AsFloat
        else VerseCIF := 0;
        Ferme(Q);
        SetControlText('MONTANT5',IntToStr(Round(verseCIF))); //PT6
        
        {if StrToFloat(GetControlText('MONTANT4')) - StrToFloat(GetControlText('MONTANT5')) > 0 then
           SetControlText('MONTANT6',FloatToStr(StrToFloat(GetControlText('MONTANT4')) - StrToFloat(GetControlText('MONTANT5')) ))
        else SetControlText('MONTANT6','0');} //PT6


        //PROFESSIONALISATION
        Q := OpenSQL('SELECT SUM(PIF_POURCENTAGE) MONTANT FROM INVESTFORMATION '+
        'WHERE PIF_ORGCOLLECTGU<>"'+OPCAInterne+'" AND PIF_MILLESIME="'+millesime+'" AND (PIF_INVESTFORM="PRO" OR PIF_INVESTFORM="DIF")',True);
        If Not Q.Eof then VersePro := Q.FindField('MONTANT').AsFloat
        else VersePro := 0;
        Ferme(Q);
        SetControlText('MONTANT7',FloatToStr(VersePro)); //PT6
        
        //SetControlText('MONTANT8',FloatToStr(StrToFloat(GetControlText('MONTANT1')) * VersePro/100)); //PT6
        Q := OpenSQL('SELECT SUM(PIF_MTVERSEMENT) MONTANT FROM INVESTFORMATION '+
        'WHERE PIF_ORGCOLLECTGU<>"'+OPCAInterne+'" AND PIF_MILLESIME="'+millesime+'" AND (PIF_INVESTFORM="PRO" OR PIF_INVESTFORM="DIF")',True);
        If Not Q.Eof then VersePro := Q.FindField('MONTANT').AsFloat
        else VersePro := 0;
        Ferme(Q);
        SetControlText('MONTANT9',IntToStr(Round(VersePro))); //PT6
        
        {If StrToFloat(GetControlText('MONTANT8')) - StrToFloat(GetControlText('MONTANT9')) > 0 then
           SetControlText('MONTANT10',FloatToStr(StrToFloat(GetControlText('MONTANT8')) - StrToFloat(GetControlText('MONTANT9'))))
        else SetControlText('MONTANT10','0');} //PT6
        
        //PLAN FORMATION
        //SetControlText('MONTANT11',FloatToStr(StrToFloat(GetControlText('MONTANT1')) * StrToFloat(GetControlText('MONTANT2'))/100)); //PT6

        // Les dépenses
        DepensesInt := 0;
        DepensesAnn := 0;
        DepensesPlu := 0;
        DepensesBdC := 0;
        DepensesExt := 0;
        //PT6 - Début
        DepExtBil   := 0;
        DepExtVAE   := 0;
        DepExtFor   := 0;
        //PT6 - Fin
        
        // Salaires des stagiaires
        Q := OpenSQL('SELECT SUM(PFO_COUTREELSAL) SALAIRESTA FROM FORMATIONS WHERE '+
        'PFO_DATEFIN>="'+UsDateTime(DateDeb)+'" AND PFO_DATEFIN<="'+UsDateTime(DateFin)+'" AND PFO_EFFECTUE="X"'+WhereMulti,True); //PT6
        If Not Q.Eof then SalaireSal := Q.FindField('SALAIRESTA').AsFloat //PT1
        else SalaireSal := 0;
        Ferme(Q);
        
        //Allocation
        Q := OpenSQL('SELECT SUM(PFO_ALLOCFORM) ALLOC FROM FORMATIONS WHERE '+
        'PFO_DATEFIN>="'+UsDateTime(DateDeb)+'" AND PFO_DATEFIN<="'+UsDateTime(DateFin)+'" AND PFO_EFFECTUE="X"'+WhereMulti,True); //PT6
        If Not Q.Eof then Alloc := Q.FindField('ALLOC').AsFloat //PT1
        else Alloc := 0;
        Ferme(Q);
        
        //Salaires des animateurs
        Q := OpenSQL('SELECT SUM(PAN_SALAIREANIM) SALAIREANIM FROM SESSIONANIMAT '+
        'LEFT JOIN SESSIONSTAGE ON PAN_CODESTAGE=PSS_CODESTAGE AND PAN_ORDRE=PSS_ORDRE AND PAN_MILLESIME=PSS_MILLESIME '+
        'WHERE PSS_EFFECTUE="X" AND '+
        'PSS_DATEFIN>="'+UsDateTime(DateDeb)+'" AND PSS_DATEFIN<="'+UsDateTime(DateFin)+'"'+WhereMultiS,True); //PT6
        If Not Q.Eof then SalaireA := Q.FindField('SALAIREANIM').AsFloat //PT1
        else SalaireA := 0;
        Ferme(Q);
        
        //Frais des stagiaires et animateurs
        Q := OpenSQL('SELECT SUM(PFS_MONTANT) FRAIS FROM FRAISSALFORM '+
        'LEFT JOIN SESSIONSTAGE ON PFS_CODESTAGE=PSS_CODESTAGE AND PFS_ORDRE=PSS_ORDRE AND PFS_MILLESIME=PSS_MILLESIME '+
        'WHERE PSS_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PSS_DATEDEBUT<="'+UsDateTime(DateFin)+'" AND PSS_EFFECTUE="X"'+WhereMultiS,True); //PT6
        If Not Q.Eof then Frais := Q.FindField('FRAIS').AsFloat //PT1
        else Frais := 0;
        Ferme(Q);
        
        //Cout pédagogique
        Q := OpenSQL('SELECT PSS_CODESTAGE,PSS_ORDRE,PSS_TYPCONVFORM,PSS_NATUREFORM,PSS_COUTPEDAG,PSS_COUTMOB,PSS_COUTFONCT,PSS_COUTORGAN,PSS_COUTUNITAIRE,PSS_ACTIONFORM,COUNT(PFO_SALARIE) NBSTA '+ //PT6
        'FROM SESSIONSTAGE LEFT JOIN FORMATIONS ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE AND PSS_MILLESIME=PFO_MILLESIME '+
        'WHERE PSS_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PSS_DATEDEBUT<="'+UsDateTime(DateFin)+'" AND PSS_EFFECTUE="X" AND PFO_EFFECTUE="X" AND PSS_INCLUSDECL="X"'+WhereMultiS+' '+ //PT6
        'GROUP BY PSS_TYPCONVFORM,PSS_NATUREFORM,PSS_COUTMOB,PSS_COUTFONCT,PSS_COUTORGAN,PSS_COUTPEDAG,PSS_COUTUNITAIRE,PSS_ACTIONFORM,PSS_CODESTAGE,PSS_ORDRE',True);  //PT1 et PT2
        TobLesDepenses := Tob.Create('LesDepenses',Nil,-1);
        TobLesDepenses.LoadDetailDB('LesDepenses','','',Q,False);
        Ferme(Q);
        
        For i := 0 to TobLesDepenses.Detail.Count - 1 do       //PT4
        begin
                CoutUnit  := TobLesDepenses.detail[i].GetValue('PSS_COUTUNITAIRE');
                NbSta     := TobLesDepenses.detail[i].GetValue('NBSTA');
                CoutPedag := TobLesDepenses.detail[i].GetValue('PSS_COUTPEDAG');
                
                If TobLesDepenses.detail[i].GetValue('PSS_NATUREFORM') = '001' then
                begin
                        DepensesInt := DepensesInt + CoutPedag + (CoutUnit * NbSta);
                end
                else
                begin
                	//PT6 - Début
                	If (TobLesDepenses.detail[i].GetValue('PSS_ACTIONFORM') <> 'BIL') And (TobLesDepenses.detail[i].GetValue('PSS_ACTIONFORM') <> 'VAE') Then
                		DepExtFor := DepExtFor + CoutPedag + (CoutUnit * NbSta)
                	Else If (TobLesDepenses.detail[i].GetValue('PSS_ACTIONFORM') = 'BIL') Then
                		DepExtBil := DepExtBil + CoutPedag + (CoutUnit * NbSta)
                	Else If (TobLesDepenses.detail[i].GetValue('PSS_ACTIONFORM') = 'VAE') Then
                		DepExtVAE := DepExtVAE + CoutPedag + (CoutUnit * NbSta);
                	//PT6 - Fin
                	
                    DepensesExt := DepensesExt + CoutPedag + (CoutUnit * NbSta);
                end;
        end;

        TobLesDepenses.Free;

        SetControlText('MONTANT11A',IntToStr(Round(DepensesInt+Frais+SalaireA))); //PT6
        If (DepensesInt+Frais+SalaireA > 0) Then SetControlVisible('BAPERCUA', True); //PT5
        //PT6 - Début
        SetControlText('MONTANT11M',IntToStr(Round(DepExtFor))); 
        SetControlText('MONTANT11N',IntToStr(Round(DepExtBil))); 
        SetControlText('MONTANT11O',IntToStr(Round(DepExtVAE))); 
        //PT6 - Fin
        SetControlText('MONTANT11B',IntToStr(Round(DepensesExt))); //PT6
        If (DepensesExt > 0) Then SetControlVisible('BAPERCUB', True); //PT5
        SetControlText('MONTANT11C',IntToStr(Round(SalaireSal))); //PT6
        SetControlText('MONTANT11D',IntToStr(Round(Alloc))); //PT6
        If (Alloc > 0) Then SetControlVisible('BAPERCUD', True); //PT5
        SetControlText('MONTANT11E',IntToStr(Round(StrToFloat(GetControlText('MONTANT5')) + StrToFloat(GetControlText('MONTANT9'))))); //PT6 Addition

        Q := OpenSQL('SELECT SUM(PIF_MTVERSEMENT) MONTANT FROM INVESTFORMATION '+
        'WHERE PIF_ORGCOLLECTGU<>"'+OPCAInterne+'" AND PIF_MILLESIME="'+millesime+'" AND PIF_INVESTFORM="DEC"',True);
        If Not Q.Eof then Versecif := Q.FindField('MONTANT').AsFloat
        else Versecif := 0;
        Ferme(Q);
        
        SetControlText('MONTANT11F',IntToStr(Round(Versecif))); //PT6
        Q := OpenSQL('SELECT SUM(PIF_MTVERSEMENT) MONTANT FROM INVESTFORMATION '+
        'WHERE PIF_ORGCOLLECTGU<>"'+OPCAInterne+'" AND PIF_MILLESIME="'+millesime+'" AND PIF_INVESTFORM="PLF"',True);
        If Not Q.Eof then VersePlan := Q.FindField('MONTANT').AsFloat
        else VersePlan := 0;
        Ferme(Q);
        
        SetControlText('MONTANT11G',IntToStr(Round(VersePlan))); //PT6
        {SetControlText('MONTANT12',FloatToStr(StrToFloat(GetControlText('MONTANT11A')) + StrToFloat(GetControlText('MONTANT11B')) + StrToFloat(GetControlText('MONTANT11C')) //PT6
        + StrToFloat(GetControlText('MONTANT11D')) + StrToFloat(GetControlText('MONTANT11E')) + StrToFloat(GetControlText('MONTANT11F')) + StrToFloat(GetControlText('MONTANT11G')) + StrToFloat(GetControlText('MONTANT11H')) )); }

        {if StrToFloat(GetControlText('MONTANT12')) - StrToFloat(GetControlText('MONTANT11')) > 0 then
           SetControlText('MONTANT13',FloatToStr(StrToFloat(GetControlText('MONTANT12')) - StrToFloat(GetControlText('MONTANT11')) ))
        else SetControlText('MONTANT13','0');
        if StrToFloat(GetControlText('MONTANT11')) - StrToFloat(GetControlText('MONTANT12')) > 0 then
           SetControlText('MONTANT14',FloatToStr(StrToFloat(GetControlText('MONTANT11')) - StrToFloat(GetControlText('MONTANT12')) ))
        else SetControlText('MONTANT14','0');} //PT6


        //MASSE SALARIALE CDD
        Q := OpenSQL('SELECT SUM(PPU_CBRUT) MASSESAL FROM PAIEENCOURS LEFT JOIN CONTRATTRAVAIL '+
        'ON PPU_SALARIE=PCI_SALARIE AND PPU_DATEDEBUT>=PCI_DEBUTCONTRAT AND PPU_DATEDEBUT<=PCI_FINCONTRAT '+
        'WHERE PCI_TYPECONTRAT="CCD" AND PPU_DATEDEBUT>="'+UsDateTime(DateDeb)+'" '+
        'AND PPU_DATEFIN<="'+UsDateTime(Datefin)+'" ',True);
        If Not Q.eof then MasseSalCDD := Q.FindField('MASSESAL').AsFloat
        else MasseSalCDD := 0;
        Ferme(Q);
        
        SetControlText('MONTANT17',IntToStr(Round(MasseSalCDD))); //PT6
        //SetControlText('MONTANT18',FloatToStr(MasseSalCDD*0.1)); //PT6
        Q := OpenSQL('SELECT SUM(PIF_MTVERSEMENT) MONTANT FROM INVESTFORMATION '+
        'WHERE PIF_ORGCOLLECTGU<>"'+OPCAInterne+'" AND PIF_MILLESIME="'+millesime+'" AND PIF_INVESTFORM="CIF"',True);
        If Not Q.Eof then VerseCIF := Q.FindField('MONTANT').AsFloat
        else VerseCIF := 0;
        Ferme(Q);
        
        SetControlText('MONTANT19',IntToStr(Round(VerseCIF))); //PT6
        {if StrToFloat(GetControlText('MONTANT18')) - StrToFloat(GetControlText('MONTANT19')) > 0 then
           SetControlText('MONTANT20',FloatToStr(StrToFloat(GetControlText('MONTANT18')) - StrToFloat(GetControlText('MONTANT19')) ))
        else SetControlText('MONTANT20','0');} //PT6
               
        // Totaux
        //PT6 - Début
        CalcTotal(GetControl('B2A'));CalcTotal(GetControl('B2B'));
        CalcTotal(GetControl('B2C'));CalcTotal(GetControl('B2D'));
        CalcTotal(GetControl('B2E'));
        CalcTotal(GetControl('B2F'));
        CalcTotal(GetControl('B2G'));
        CalcTotal(GetControl('B2H'));
        
        CalcMontant(GetControl('MONTANT1'));
        CalcMontant(GetControl('MONTANT5'));
        CalcMontant(GetControl('MONTANT11A'));
        CalcMontant(GetControl('MONTANT15'));
        CalcMontant(GetControl('MONTANT17'));
        //PT6 - Fin
        
        //PT6 - Début
        // Reports année précédente
        MillesimeOld := IntToStr((StrToInt(Millesime) - 1));
        Q := OpenSQL('SELECT PFD_I2A,PFD_I3C,PFD_I2B,PFD_I3D,PFD_I4E,PFD_F13,PFD_F14 FROM DECLFORMATION2483 WHERE PFD_MILLESIME="'+MillesimeOld+'"', True);
        If Not Q.EOF Then
        Begin
        	ExcA := Round(Q.FindField('PFD_I2B').AsFloat - Q.FindField('PFD_I3D').AsFloat);
        	SetControlText('EXCEDENTA', FloatToStr(ExcA));
        	ExcB := Q.FindField('PFD_I4E').AsFloat;
        	SetControlText('EXCEDENTB', FloatToStr(ExcB));
        	ExcC := Round(Q.FindField('PFD_F14').AsFloat - (Q.FindField('PFD_I2A').AsFloat - Q.FindField('PFD_I3C').AsFloat));
        	If ExcC < 0 Then ExcC := 0; If ExcC > ExcA Then ExcC := ExcA;
        	SetControlText('EXCEDENTC', FloatToStr(ExcC));
        	ExcD := Round(Q.FindField('PFD_F14').AsFloat -
        		((Q.FindField('PFD_I2A').AsFloat - Q.FindField('PFD_I3C').AsFloat) + (Q.FindField('PFD_I2B').AsFloat - Q.FindField('PFD_I3D').AsFloat))
        		);
        	If ExcD < 0 Then ExcD := 0; If ExcD > ExcB Then ExcD := ExcB;
        	SetControlText('EXCEDENTD', FloatToStr(ExcD));
        	ExcE := Q.FindField('PFD_F13').AsFloat;
        	SetControlText('EXCEDENTE', FloatToStr(ExcE));
    	End;
    	
    	//Totaux reports des années précédentes
    	CalcExcedents(GetControl('EXCEDENTA'));
    	CalcExcedents(GetControl('EXCEDENTB'));
        //PT6 - Fin
end;

procedure TOF_PGDECLARAFORMATION.EditDeclaration(Sender: TObject);
var TobEtat,T : Tob;
    Pages : TPageControl;
    StPages : String;
begin
        Pages := TPageControl(GetControl('PAGE1_'));
        StPages := '';
        //Nouvelle édition
        TobEtat := Tob.Create('TobEdition',Nil,-1);
        T := Tob.Create('Fille',TobEtat,-1);
        T.AddChampSupValeur('JOURHEURE',GetControlText('JOURHEURE'));
        T.AddChampSupValeur('ADRESSE1',GetControlText('ADRESSE1'));
        T.AddChampSupValeur('ADRESSE2',GetControlText('ADRESSE2'));
        T.AddChampSupValeur('ADRESSE3',GetControlText('ADRESSE3'));
        T.AddChampSupValeur('IDENTIFIANT',GetControlText('IDENTIFIANT'));
        T.AddChampSupValeur('ADRESSEDECL1',GetControlText('ADRESSEDECL1'));
        T.AddChampSupValeur('ADRESSEDECL2',GetControlText('ADRESSEDECL2'));
        T.AddChampSupValeur('ADRESSEDECL3',GetControlText('ADRESSEDECL3'));
        T.AddChampSupValeur('SIRET1',GetControlText('SIRET1'));
        T.AddChampSupValeur('SIRET2',GetControlText('SIRET2'));
        T.AddChampSupValeur('ACTIVITE',GetControlText('ACTIVITE'));
        T.AddChampSupValeur('CHANGEMENT1',GetControlText('CHANGEMENT1'));
        T.AddChampSupValeur('CHANGEMENT2',GetControlText('CHANGEMENT2'));
        T.AddChampSupValeur('CHANGEMENT3',GetControlText('CHANGEMENT3'));
        T.AddChampSupValeur('LIEU',GetControlText('LIEU'));
        T.AddChampSupValeur('DATE',GetControlText('DATE'));
        T.AddChampSupValeur('HONNEUR',GetControlText('HONNEUR'));
        T.AddChampSupValeur('JUSTIFIER',GetControlText('JUSTIFIER'));

    	If StrToInt(Millesime) >= 2007 Then //PT6
            LanceEtatTOB('E','PFO','PX'+RightStr(Millesime,1),TobEtat,True,False,False,Pages,'','',False,0,StPages)
        Else
            LanceEtatTOB('E','PFO','PFD',TobEtat,True,False,False,Pages,'','',False,0,StPages);
        TobEtat.Free;

        TobEtat := Tob.Create('TobEdition',Nil,-1);
        T := Tob.Create('Fille',TobEtat,-1);
        T.AddChampSupValeur('TOTAL',GetControlText('TOTAL'));

    	If StrToInt(Millesime) >= 2007 Then //PT6
            LanceEtatTOB('E','PFO','PY'+RightStr(Millesime,1),TobEtat,True,False,False,Pages,'','',False,0,StPages)
        Else
            LanceEtatTOB('E','PFO','PFE',TobEtat,True,False,False,Pages,'','',False,0,StPages);
        TobEtat.Free;
end;

procedure TOF_PGDECLARAFORMATION.OnArgument(Arguments : String ) ;
var 
    BCalc,BImp,BVal : TToolBarButton97;
    MenuPop : TPopUpMenu; //PT5
    i : Integer; //PT5
    Edit : THEdit; //PT6
begin
inherited ;
	Millesime := ReadTokenPipe(Arguments,';');
	BCalc := TToolBarButton97(getControl('BCALCULER'));
	If BCalc <> Nil then BCalc.OnClick := CalculerDecl;
	Bimp := TToolBarButton97(GetControl('BImprimer'));
	If BImp <> Nil then BImp.OnClick :=  EditDeclaration;
	BVal := TToolBarButton97(GetControl('BValider'));
	If BVal <> Nil then BVal.OnClick :=  ValiderDecl;
	
	// Boutons Aperçu
	//PT5 - Début
	MenuPop := TPopUpMenu(GetControl('POPUPVISUA'));
	If MenuPop <> Nil then
	Begin
		For i := 0 to MenuPop.Items.Count - 1 do
		Begin
			MenuPop.Items[i].OnClick := EditDetailsA;
		End;
	End;
	
	MenuPop := TPopUpMenu(GetControl('POPUPVISUB'));
	If MenuPop <> Nil then
	Begin
		For i := 0 to MenuPop.Items.Count - 1 do
		Begin
			MenuPop.Items[i].OnClick := EditDetailsB;
		End;
	End;
	
	MenuPop := TPopUpMenu(GetControl('POPUPVISUD'));
	If MenuPop <> Nil then
	Begin
		For i := 0 to MenuPop.Items.Count - 1 do
		Begin
			MenuPop.Items[i].OnClick := EditDetailsD;
		End;
	End;
	//PT5 - Fin
	
	ChargerDecl;
	SetControlText('SIRET1',Copy(GetParamSocSecur(('SO_SIRET'),''),1,9));
	SetControlText('SIRET2',Copy(GetParamSocSecur(('SO_SIRET'),''),10,5));
	SetControlText('ACTIVITE', GetParamSocSecur(('SO_APE'),''));
	SetControlText('IDENTIFIANT', GetParamSocSecur(('SO_LIBELLE'),''));
	SetControlText('ADRESSEDECL1', GetParamSocSecur(('SO_ADRESSE1'),''));
	SetControlText('ADRESSEDECL2', GetParamSocSecur(('SO_ADRESSE2'),''));
	SetControlText('ADRESSEDECL3', GetParamSocSecur(('SO_CODEPOSTAL'),'') + ' ' + GetParamSocSecur(('SO_VILLE'),''));
	
	//PT6 - Début
	// Gestion des totaux et cumuls : Si modification d'un champ, recalcul automatique des sommes
	Edit := THEdit(GetControl('B2A')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B3A')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B4A')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B5A')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	
	Edit := THEdit(GetControl('B2B')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B3B')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B4B')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B5B')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	
	Edit := THEdit(GetControl('B2C')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B3C')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B4C')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B5C')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	
	Edit := THEdit(GetControl('B2D')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B3D')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B4D')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B5D')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	
	Edit := THEdit(GetControl('B2E')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B3E')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B4E')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B5E')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	
	Edit := THEdit(GetControl('B2F')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B3F')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B4F')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B5F')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	
	Edit := THEdit(GetControl('B2G')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B3G')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B4G')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B5G')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	
	Edit := THEdit(GetControl('B2H')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B3H')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B4H')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	Edit := THEdit(GetControl('B5H')); If Edit <> Nil Then Edit.OnExit := CalcTotal;
	
	Edit := THEdit(GetControl('MONTANT1'));   If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT2'));   If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT3'));   If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT5'));   If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT7'));   If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT9'));   If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT11'));  If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT11A')); If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT11B')); If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT11C')); If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT11D')); If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT11E')); If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT11F')); If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT11G')); If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT11H')); If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT11I')); If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT11M')); If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT11N')); If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT11O')); If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT15'));  If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT17'));  If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT19'));  If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT22'));  If Edit <> Nil Then Edit.OnExit := CalcMontant;
	Edit := THEdit(GetControl('MONTANT23'));  If Edit <> Nil Then Edit.OnExit := CalcMontant;
	
	Edit := THEdit(GetControl('EXCEDENTA'));  If Edit <> Nil Then Edit.OnExit := CalcExcedents;
	Edit := THEdit(GetControl('EXCEDENTB'));  If Edit <> Nil Then Edit.OnExit := CalcExcedents;
	Edit := THEdit(GetControl('EXCEDENTC'));  If Edit <> Nil Then Edit.OnExit := CalcExcedents;
	Edit := THEdit(GetControl('EXCEDENTD'));  If Edit <> Nil Then Edit.OnExit := CalcExcedents;
	Edit := THEdit(GetControl('EXCEDENTE'));  If Edit <> Nil Then Edit.OnExit := CalcExcedents;
	
	
	// Changement de fond page pour garder l'historique de visualisation
	// Pour les années antérieures à 2007, on ne faisait pas d'historique donc on garde le fond par défaut
	If StrToInt(Millesime) >= 2007 Then
	Begin
		TPDFPreview(GetControl('PDF1')).PDFPath := '$GED=PAIE;24831_'+Millesime+'.pdf;PDF;;;;;-;FRA;CEG';
		TPDFPreview(GetControl('PDF2')).PDFPath := '$GED=PAIE;24832_'+Millesime+'.pdf;PDF;;;;;-;FRA;CEG';
        SetControlVisible('ACTIVITE', False);
	End;
	//PT6 - Fin
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 11/02/2008 / PT6
Modifié le ... :   /  /
Description .. : Calcul des totaux de la page 1
Mots clefs ... :
*****************************************************************}
procedure TOF_PGDECLARAFORMATION.CalcTotal (Sender : TObject);
Var ChampEnCours, Col : String;
    Somme 			  : Double;
    i 				  : Integer;
Begin
	ChampEnCours := THEdit(Sender).Name;
	If LeftStr(ChampEnCours,1) <> 'B' Then Exit;
	
	Col := RightStr(ChampEnCours,1);
	Somme := 0;
	For i := 2 To 5 Do
	Begin
		Somme := Somme + StrToFloat(GetControlText('B'+IntToStr(i)+Col));
	End;
	
	SetControlText('B6' + Col, FloatToStr(Somme));
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 14/04/2008 / PT6
Modifié le ... :   /  /
Description .. : Calcul des excedents
Mots clefs ... :
*****************************************************************}
procedure TOF_PGDECLARAFORMATION.CalcExcedents (Sender : TObject);
Var ChampEnCours : String;
	Somme        : Integer;
Begin
	ChampEnCours := THEdit(Sender).Name;
	
	If (RightStr(ChampEnCours,1) = 'A') Or (RightStr(ChampEnCours,1) = 'C') Then
		SetControlText('TOTEXCEDENT1', IntToStr(Round(
		StrToFloat(GetControlText('EXCEDENTA')) - StrToFloat(GetControlText('EXCEDENTC'))
		)));

	If (RightStr(ChampEnCours,1) = 'B') Or (RightStr(ChampEnCours,1) = 'D') Then
		SetControlText('TOTEXCEDENT2', IntToStr(Round(
		StrToFloat(GetControlText('EXCEDENTB')) - StrToFloat(GetControlText('EXCEDENTD'))
		)));
	
	Somme := StrToInt(GetControlText('TOTEXCEDENT1')) + StrToInt(GetControlText('TOTEXCEDENT2')) + StrToInt(GetControlText('EXCEDENTE'));
	SetControlText('TOTAL', IntToStr(Somme));
	SetControlText('MONTANT15', IntToStr(Somme));
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 11/02/2008 / PT6
Modifié le ... :   /  /
Description .. : Calcul des totaux de la page 2
Mots clefs ... :
*****************************************************************}
procedure TOF_PGDECLARAFORMATION.CalcMontant (Sender : TObject);
Var ChampEnCours : String;
    Somme 		 : Double;
    i,NumChamp	 : Integer;
    c : Char;
    
    // Sous-Fonction de calcul de la ligne 4
    Function CalcLigne4 () : Integer;
    Begin
			Somme := StrToFloat(GetControlText('MONTANT1')) * StrToFloat(GetControlText('MONTANT3'));
			Result:= Round(Somme);
    End;
    // Sous-Fonction de calcul de la ligne 6
    Function CalcLigne6 () : Integer;
    Begin
			Somme := StrToFloat(GetControlText('MONTANT4')) - StrToFloat(GetControlText('MONTANT5'));
			i     := Round(Somme);
			If i > 0 Then Result := i Else Result := 0;
    End;
    // Sous-Fonction de calcul de la ligne 8
    Function CalcLigne8 () : Integer;
    Begin
			Somme := StrToFloat(GetControlText('MONTANT1')) * StrToFloat(GetControlText('MONTANT7'));
			Result:= Round(Somme);
    End;
    // Sous-Fonction de calcul de la ligne 10
    Function CalcLigne10 () : Integer;
    Begin
			Somme := StrToFloat(GetControlText('MONTANT8')) - StrToFloat(GetControlText('MONTANT9'));
			i     := Round(Somme);
			If i > 0 Then Result := i Else Result := 0;
    End;
    // Sous-Fonction de calcul de la ligne 11
    Function CalcLigne11 () : Integer;
    Begin
			Somme := StrToFloat(GetControlText('MONTANT1')) * StrToFloat(GetControlText('MONTANT2'));
			Result:= Round(Somme);
    End;
    // Sous-Fonction de calcul de la ligne 11B
    Function CalcLigne11B () : Integer;
    Begin
			Somme := StrToFloat(GetControlText('MONTANT11M')) + StrToFloat(GetControlText('MONTANT11N')) + StrToFloat(GetControlText('MONTANT11O'));
			Result:= Round(Somme);
    End;
    // Sous-Fonction de calcul de la ligne 11E
    Function CalcLigne11E () : Integer;
    Begin
			Somme := StrToFloat(GetControlText('MONTANT5')) + StrToFloat(GetControlText('MONTANT9'));
			Result:= Round(Somme);
    End;
    // Sous-Fonction de calcul de la ligne 12
    Function CalcLigne12 () : Integer;
    Begin
			Somme := StrToFloat(GetControlText('MONTANT11A')) + StrToFloat(GetControlText('MONTANT11B')) +
					 StrToFloat(GetControlText('MONTANT11C')) + StrToFloat(GetControlText('MONTANT11D')) +
					 StrToFloat(GetControlText('MONTANT11E')) + StrToFloat(GetControlText('MONTANT11F')) +
					 StrToFloat(GetControlText('MONTANT11G')) + StrToFloat(GetControlText('MONTANT11H')) -
					 StrToFloat(GetControlText('MONTANT11I'));
			Result:= Round(Somme);
    End;
    // Sous-Fonction de calcul de la ligne 13
    Function CalcLigne13 () : Integer;
    Begin
			Somme := StrToFloat(GetControlText('MONTANT12')) - StrToFloat(GetControlText('MONTANT11'));
			i     := Round(Somme);
			If i > 0 Then Result := i Else Result := 0;
    End;
    // Sous-Fonction de calcul de la ligne 14
    Function CalcLigne14 () : Integer;
    Begin
			Somme := StrToFloat(GetControlText('MONTANT11')) - StrToFloat(GetControlText('MONTANT12'));
			i     := Round(Somme);
			If i > 0 Then Result := i Else Result := 0;
    End;
    // Sous-Fonction de calcul de la ligne 16
    Function CalcLigne16 () : Integer;
    Begin
			Somme := StrToFloat(GetControlText('MONTANT14')) - StrToFloat(GetControlText('MONTANT15'));
			i     := Round(Somme);
			If i > 0 Then Result := i Else Result := 0;
    End;
    // Sous-Fonction de calcul de la ligne 18
    Function CalcLigne18 () : Integer;
    Begin
			Somme := StrToFloat(GetControlText('MONTANT17')) / 100;
			Result:= Round(Somme);
    End;
    // Sous-Fonction de calcul de la ligne 20
    Function CalcLigne20 () : Integer;
    Begin
			Somme := StrToFloat(GetControlText('MONTANT18')) - StrToFloat(GetControlText('MONTANT19'));
			i     := Round(Somme);
			If i > 0 Then Result := i Else Result := 0;
    End;
    // Sous-Fonction de calcul de la ligne 24
    Function CalcLigne24 () : Integer;
    Begin
			Somme := StrToFloat(GetControlText('MONTANT6')) + StrToFloat(GetControlText('MONTANT16')) +
					(StrToFloat(GetControlText('MONTANT20')) * 2);
			Result:= Round(Somme);
    End;
    // Sous-Fonction de calcul de la ligne 25
    Function CalcLigne25 () : Integer;
    Begin
			Somme := StrToFloat(GetControlText('MONTANT21')) + StrToFloat(GetControlText('MONTANT22')) +
					 StrToFloat(GetControlText('MONTANT23')) + StrToFloat(GetControlText('MONTANT24'));
			Result:= Round(Somme);
    End;
    // Sous-Fonction de calcul des lignes 11
    Procedure CalcLigne11X;
    Begin
    		SetControlText('MONTANT11B', IntToStr(CalcLigne11B()));
			SetControlText('MONTANT12', IntToStr(CalcLigne12()));
			SetControlText('MONTANT13', IntToStr(CalcLigne13()));
			SetControlText('MONTANT14', IntToStr(CalcLigne14()));
			SetControlText('MONTANT16', IntToStr(CalcLigne16()));
			SetControlText('MONTANT24', IntToStr(CalcLigne24()));
			SetControlText('MONTANT25', IntToStr(CalcLigne25()));
    End;
Begin
	ChampEnCours := THEdit(Sender).Name;
	If LeftStr(ChampEnCours,7) <> 'MONTANT' Then Exit;
	
	If IsNumeric(Copy(ChampEnCours, 8, Length(ChampEnCours))) Then
		NumChamp := StrToInt(Copy(ChampEnCours, 8, Length(ChampEnCours)))
	Else
    Begin
		// Cas particuliers des champs 11a à 11i
        c := ChampEnCours[Length(ChampEnCours)];
		NumChamp := Ord(c);
    End;
	
	Case NumChamp Of
		1 :
		  Begin
			SetControlText('MONTANT4', IntToStr(CalcLigne4()));
			SetControlText('MONTANT6', IntToStr(CalcLigne6()));
			SetControlText('MONTANT8', IntToStr(CalcLigne8()));
			SetControlText('MONTANT10', IntToStr(CalcLigne10()));
			SetControlText('MONTANT11', IntToStr(CalcLigne11()));
		  	SetControlText('MONTANT13', IntToStr(CalcLigne13()));
		  	SetControlText('MONTANT14', IntToStr(CalcLigne14()));
		  	SetControlText('MONTANT16', IntToStr(CalcLigne16()));
		  	SetControlText('MONTANT21', IntToStr(CalcLigne10()));
		  	SetControlText('MONTANT24', IntToStr(CalcLigne24()));
		  	SetControlText('MONTANT25', IntToStr(CalcLigne25()));
		  End;
		2 :
		  Begin
			SetControlText('MONTANT11', IntToStr(CalcLigne11()));
		  	SetControlText('MONTANT13', IntToStr(CalcLigne13()));
		  	SetControlText('MONTANT14', IntToStr(CalcLigne14()));
		  	SetControlText('MONTANT16', IntToStr(CalcLigne16()));
		  	SetControlText('MONTANT24', IntToStr(CalcLigne24()));
		  	SetControlText('MONTANT25', IntToStr(CalcLigne25()));
		  End;
		3 :
		  Begin
			SetControlText('MONTANT4', IntToStr(CalcLigne4()));
			SetControlText('MONTANT6', IntToStr(CalcLigne6()));
			SetControlText('MONTANT24', IntToStr(CalcLigne24()));
			SetControlText('MONTANT25', IntToStr(CalcLigne25()));
		  End;
		5 :
		  Begin
			SetControlText('MONTANT6', IntToStr(CalcLigne6()));
			SetControlText('MONTANT11E', IntToStr(CalcLigne11E()));
			SetControlText('MONTANT24', IntToStr(CalcLigne24()));
			SetControlText('MONTANT25', IntToStr(CalcLigne25()));
		  End;
		7 :
		  Begin
			SetControlText('MONTANT8', IntToStr(CalcLigne8()));
			SetControlText('MONTANT10', IntToStr(CalcLigne10()));
			SetControlText('MONTANT21', IntToStr(CalcLigne10()));
			SetControlText('MONTANT25', IntToStr(CalcLigne25()));
		  End;
		9 :
		  Begin
		  	SetControlText('MONTANT10', IntToStr(CalcLigne10()));
			SetControlText('MONTANT11E', IntToStr(CalcLigne11E()));
			SetControlText('MONTANT21', IntToStr(CalcLigne10()));
			SetControlText('MONTANT25', IntToStr(CalcLigne25()));
		  End;
		65 : //11A
			CalcLigne11X;
		66 : //11B
		  	CalcLigne11X;
		67 : //11C
		  	CalcLigne11X;
		68 : //11D
		  	CalcLigne11X;
		69 : //11E
		  	CalcLigne11X;
		70 : //11F
		  	CalcLigne11X;
		71 : //11G
		  	CalcLigne11X;
		72 : //11H
		  	CalcLigne11X;
		73 : //11I
		  	CalcLigne11X;
		77 : //11M
		  	CalcLigne11X;
		78 : //11N
		  	CalcLigne11X;
		79 : //11O
		  	CalcLigne11X;
		15 :
		  Begin
		  	SetControlText('MONTANT16', IntToStr(CalcLigne16()));
		  	SetControlText('MONTANT24', IntToStr(CalcLigne24()));
		  	SetControlText('MONTANT25', IntToStr(CalcLigne25()));
		  End;
		17 :
		  Begin
		  	SetControlText('MONTANT18', IntToStr(CalcLigne18()));
		  	SetControlText('MONTANT20', IntToStr(CalcLigne20()));
		  	SetControlText('MONTANT24', IntToStr(CalcLigne24()));
		  	SetControlText('MONTANT25', IntToStr(CalcLigne25()));
		  End;
		19 :
		  Begin
		  	SetControlText('MONTANT20', IntToStr(CalcLigne20()));
		  	SetControlText('MONTANT24', IntToStr(CalcLigne24()));
		  	SetControlText('MONTANT25', IntToStr(CalcLigne25()));
		  End;
		22 :
		  Begin
		  	SetControlText('MONTANT25', IntToStr(CalcLigne25()));
		  End;
		23 :
		  Begin
		  	SetControlText('MONTANT25', IntToStr(CalcLigne25()));
		  End;
	End;
End;

procedure TOF_PGDECLARAFORMATION.ValiderDecl(Sender : TObject);
var TobDecl,T : Tob;
begin
     TobDecl := Tob.Create('DECLFORMATION2483',Nil,-1);
     TobDecl.LoadDetailDB('DECLFORMATION2483', '', '', Nil, False);
     T:= TobDecl.FindFirst(['PFD_MILLESIME'],[Millesime],False);
     If T = Nil then
     begin
      T := Tob.Create('DECLFORMATION2483',TobDecl,-1);
      T.PutValue('PFD_MILLESIME',Millesime);
     end;
     T.PutValue('PFD_NBSALMOYEN',StrToFloat(GetControlText('NBSALARIES')));
     T.PutValue('PFD_NBOUV1',StrToFloat(GetControlText('B2A')));
     T.PutValue('PFD_NBOUV2',StrToFloat(GetControlText('B2B')));
     T.PutValue('PFD_NBOUV3',StrToFloat(GetControlText('B2C')));
     T.PutValue('PFD_NBOUV4',StrToFloat(GetControlText('B2D')));
     T.PutValue('PFD_NBOUV5',StrToFloat(GetControlText('B2E')));
     T.PutValue('PFD_NBOUV6',StrToFloat(GetControlText('B2F')));
     T.PutValue('PFD_NBOUV7',StrToFloat(GetControlText('B2G')));
     T.PutValue('PFD_NBOUV8',StrToFloat(GetControlText('B2H')));
     T.PutValue('PFD_NBEMP1',StrToFloat(GetControlText('B3A')));
     T.PutValue('PFD_NBEMP2',StrToFloat(GetControlText('B3B')));
     T.PutValue('PFD_NBEMP3',StrToFloat(GetControlText('B3C')));
     T.PutValue('PFD_NBEMP4',StrToFloat(GetControlText('B3D')));
     T.PutValue('PFD_NBEMP5',StrToFloat(GetControlText('B3E')));
     T.PutValue('PFD_NBEMP6',StrToFloat(GetControlText('B3F')));
     T.PutValue('PFD_NBEMP7',StrToFloat(GetControlText('B3G')));
     T.PutValue('PFD_NBEMP8',StrToFloat(GetControlText('B3H')));
     T.PutValue('PFD_NBTEC1',StrToFloat(GetControlText('B4A')));
     T.PutValue('PFD_NBTEC2',StrToFloat(GetControlText('B4B')));
     T.PutValue('PFD_NBTEC3',StrToFloat(GetControlText('B4C')));
     T.PutValue('PFD_NBTEC4',StrToFloat(GetControlText('B4D')));
     T.PutValue('PFD_NBTEC5',StrToFloat(GetControlText('B4E')));
     T.PutValue('PFD_NBTEC6',StrToFloat(GetControlText('B4F')));
     T.PutValue('PFD_NBTEC7',StrToFloat(GetControlText('B4G')));
     T.PutValue('PFD_NBTEC8',StrToFloat(GetControlText('B4H')));
     T.PutValue('PFD_NBCAD1',StrToFloat(GetControlText('B5A')));
     T.PutValue('PFD_NBCAD2',StrToFloat(GetControlText('B5B')));
     T.PutValue('PFD_NBCAD3',StrToFloat(GetControlText('B5C')));
     T.PutValue('PFD_NBCAD4',StrToFloat(GetControlText('B5D')));
     T.PutValue('PFD_NBCAD5',StrToFloat(GetControlText('B5E')));
     T.PutValue('PFD_NBCAD6',StrToFloat(GetControlText('B5F')));
     T.PutValue('PFD_NBCAD7',StrToFloat(GetControlText('B5G')));
     T.PutValue('PFD_NBCAD8',StrToFloat(GetControlText('B5H')));
     T.PutValue('PFD_NBTOT1',StrToFloat(GetControlText('B6A')));
     T.PutValue('PFD_NBTOT2',StrToFloat(GetControlText('B6B')));
     T.PutValue('PFD_NBTOT3',StrToFloat(GetControlText('B6C')));
     T.PutValue('PFD_NBTOT4',StrToFloat(GetControlText('B6D')));
     T.PutValue('PFD_NBTOT5',StrToFloat(GetControlText('B6E')));
     T.PutValue('PFD_NBTOT6',StrToFloat(GetControlText('B6F')));
     T.PutValue('PFD_NBTOT7',StrToFloat(GetControlText('B6G')));
     T.PutValue('PFD_NBTOT8',StrToFloat(GetControlText('B6H')));
     T.PutValue('PFD_A7',StrToFloat(GetControlText('B7')));
     T.PutValue('PFD_A8',StrToFloat(GetControlText('B8')));
     T.PutValue('PFD_A9',StrToFloat(GetControlText('B9')));
     T.PutValue('PFD_A10',StrToFloat(GetControlText('B10')));
     T.PutValue('PFD_A11',StrToFloat(GetControlText('B11')));
     T.PutValue('PFD_A12',StrToFloat(GetControlText('B12')));
     T.PutValue('PFD_C1',StrToFloat(GetControlText('MONTANT1')));
     T.PutValue('PFD_C2',StrToFloat(GetControlText('MONTANT2')));
//     T.PutValue('PFD_C3',StrToFloat(GetControlText('MONTANT3')));
 //    T.PutValue('PFD_C4',StrToFloat(GetControlText('MONTANT4')));
//     T.PutValue('PFD_C5',StrToFloat(GetControlText('MONTANT5')));
//     T.PutValue('PFD_C6',StrToFloat(GetControlText('MONTANT6')));
     T.PutValue('PFD_D3',StrToFloat(GetControlText('MONTANT3')));
     T.PutValue('PFD_D4',StrToFloat(GetControlText('MONTANT4')));
     T.PutValue('PFD_D5',StrToFloat(GetControlText('MONTANT5')));
     T.PutValue('PFD_D6',StrToFloat(GetControlText('MONTANT6')));
     T.PutValue('PFD_E7',StrToFloat(GetControlText('MONTANT7')));
     T.PutValue('PFD_E8',StrToFloat(GetControlText('MONTANT8')));
     T.PutValue('PFD_E9',StrToFloat(GetControlText('MONTANT9')));
     T.PutValue('PFD_F10',StrToFloat(GetControlText('MONTANT11')));
     T.PutValue('PFD_F10A',StrToFloat(GetControlText('MONTANT11A')));
     T.PutValue('PFD_F10B',StrToFloat(GetControlText('MONTANT11B')));
     T.PutValue('PFD_F10C',StrToFloat(GetControlText('MONTANT11C')));
     T.PutValue('PFD_F10D',StrToFloat(GetControlText('MONTANT11D')));
     T.PutValue('PFD_F10E',StrToFloat(GetControlText('MONTANT11E')));
     T.PutValue('PFD_F10F',StrToFloat(GetControlText('MONTANT11F')));
     T.PutValue('PFD_F10G',StrToFloat(GetControlText('MONTANT11G')));
     T.PutValue('PFD_F10H',StrToFloat(GetControlText('MONTANT11H')));
     T.PutValue('PFD_F11',StrToFloat(GetControlText('MONTANT11I')));
     T.PutValue('PFD_F12',StrToFloat(GetControlText('MONTANT12')));
     T.PutValue('PFD_F13',StrToFloat(GetControlText('MONTANT13')));
     T.PutValue('PFD_F14',StrToFloat(GetControlText('MONTANT14')));
     T.PutValue('PFD_F15',StrToFloat(GetControlText('MONTANT15')));
     T.PutValue('PFD_G16',StrToFloat(GetControlText('MONTANT16')));
     T.PutValue('PFD_G17',StrToFloat(GetControlText('MONTANT17')));
     T.PutValue('PFD_G18',StrToFloat(GetControlText('MONTANT18')));
     T.PutValue('PFD_G19',StrToFloat(GetControlText('MONTANT19')));
     T.PutValue('PFD_H20',StrToFloat(GetControlText('MONTANT20')));
     T.PutValue('PFD_H21',StrToFloat(GetControlText('MONTANT21')));
     T.PutValue('PFD_H22',StrToFloat(GetControlText('MONTANT22')));
     T.PutValue('PFD_H23',StrToFloat(GetControlText('MONTANT23')));
     T.PutValue('PFD_H24',StrToFloat(GetControlText('MONTANT24')));
     T.PutValue('PFD_I2A',StrToFloat(GetControlText('EXCEDENTA')));
     T.PutValue('PFD_I2B',StrToFloat(GetControlText('EXCEDENTB')));
     T.PutValue('PFD_I3C',StrToFloat(GetControlText('EXCEDENTC')));
     T.PutValue('PFD_I3D',StrToFloat(GetControlText('EXCEDENTD')));
     T.PutValue('PFD_I4E',StrToFloat(GetControlText('EXCEDENTE')));
     T.PutValue('PFD_CEDITIMPOT1',StrToFloat(GetControlText('TOTAL')));
     T.PutValue('PFD_ATTESTHON',GetControlText('HONNEUR'));
     T.PutValue('PFD_OBLIG',GetControlText('JUSTIFIER'));
     T.InsertOrUpdateDB(False);
     TobDecl.Free;
end;

procedure TOF_PGDECLARAFORMATION.ChargerDecl;
Var Q : TQuery;
begin
     Q := OpenSQL('SELECT * FROM DECLFORMATION2483 WHERE PFD_MILLESIME="'+Millesime+'"',True);
     If Not Q.Eof then
     begin
       SetControlText('NBSALARIES',Q.FindField('PFD_NBSALMOYEN').AsString);
       SetControlText('B2A',Q.FindField('PFD_NBOUV1').AsString);
       SetControlText('B2B',Q.FindField('PFD_NBOUV2').AsString);
       SetControltext('B2C',Q.FindField('PFD_NBOUV3').AsString);
       SetControltext('B2D',Q.FindField('PFD_NBOUV4').AsString);
       SetControltext('B2E',Q.FindField('PFD_NBOUV5').AsString);
       SetControltext('B2F',Q.FindField('PFD_NBOUV6').AsString);
       SetControltext('B2G',Q.FindField('PFD_NBOUV7').AsString);
       SetControltext('B2H',Q.FindField('PFD_NBOUV8').AsString);
       SetControltext('B3A',Q.FindField('PFD_NBEMP1').AsString);
       SetControltext('B3B',Q.FindField('PFD_NBEMP2').AsString);
       SetControltext('B3C',Q.FindField('PFD_NBEMP3').AsString);
       SetControltext('B3D',Q.FindField('PFD_NBEMP4').AsString);
       SetControltext('B3E',Q.FindField('PFD_NBEMP5').AsString);
       SetControltext('B3F',Q.FindField('PFD_NBEMP6').AsString);
       SetControltext('B3G',Q.FindField('PFD_NBEMP7').AsString);
       SetControltext('B3H',Q.FindField('PFD_NBEMP8').AsString);
       SetControltext('B4A',Q.FindField('PFD_NBTEC1').AsString);
       SetControltext('B4B',Q.FindField('PFD_NBTEC2').AsString);
       SetControltext('B4C',Q.FindField('PFD_NBTEC3').AsString);
       SetControltext('B4D',Q.FindField('PFD_NBTEC4').AsString);
       SetControltext('B4E',Q.FindField('PFD_NBTEC5').AsString);
       SetControltext('B4F',Q.FindField('PFD_NBTEC6').AsString);
       SetControltext('B4G',Q.FindField('PFD_NBTEC7').AsString);
       SetControltext('B4H',Q.FindField('PFD_NBTEC8').AsString);
       SetControltext('B5A',Q.FindField('PFD_NBCAD1').AsString);
       SetControltext('B5B',Q.FindField('PFD_NBCAD2').AsString);
       SetControltext('B5C',Q.FindField('PFD_NBCAD3').AsString);
       SetControltext('B5D',Q.FindField('PFD_NBCAD4').AsString);
       SetControltext('B5E',Q.FindField('PFD_NBCAD5').AsString);
       SetControltext('B5F',Q.FindField('PFD_NBCAD6').AsString);
       SetControltext('B5G',Q.FindField('PFD_NBCAD7').AsString);
       SetControltext('B5H',Q.FindField('PFD_NBCAD8').AsString);
       SetControltext('B6A',Q.FindField('PFD_NBTOT1').AsString);
       SetControltext('B6B',Q.FindField('PFD_NBTOT2').AsString);
       SetControltext('B6C',Q.FindField('PFD_NBTOT3').AsString);
       SetControltext('B6D',Q.FindField('PFD_NBTOT4').AsString);
       SetControltext('B6E',Q.FindField('PFD_NBTOT5').AsString);
       SetControltext('B6F',Q.FindField('PFD_NBTOT6').AsString);
       SetControltext('B6G',Q.FindField('PFD_NBTOT7').AsString);
       SetControltext('B6H',Q.FindField('PFD_NBTOT8').AsString);
       SetControltext('B7',Q.FindField('PFD_A7').AsString);
       SetControltext('B8',Q.FindField('PFD_A8').AsString);
       SetControltext('B9',Q.FindField('PFD_A9').AsString);
       SetControltext('B10',Q.FindField('PFD_A10').AsString);
       SetControltext('B11',Q.FindField('PFD_A11').AsString);
       SetControltext('B12',Q.FindField('PFD_A12').AsString);
       SetControltext('MONTANT1',Q.FindField('PFD_C1').AsString);
       SetControltext('MONTANT2',Q.FindField('PFD_C2').AsString);
  //     SetControltext('',Q.FindField('PFD_C3',FloatToStr(SetControlText('MONTANT3')));
   //    SetControltext('',Q.FindField('PFD_C4',FloatToStr(SetControlText('MONTANT4')));
  //     SetControltext('',Q.FindField('PFD_C5',FloatToStr(SetControlText('MONTANT5')));
  //     SetControltext('',Q.FindField('PFD_C6',FloatToStr(SetControlText('MONTANT6')));
       SetControltext('MONTANT3',Q.FindField('PFD_D3').AsString);
       SetControltext('MONTANT4',Q.FindField('PFD_D4').AsString);
       SetControltext('MONTANT5',Q.FindField('PFD_D5').AsString);
       SetControltext('MONTANT6',Q.FindField('PFD_D6').AsString);
       SetControltext('MONTANT7',Q.FindField('PFD_E7').AsString);
       SetControltext('MONTANT8',Q.FindField('PFD_E8').AsString);
       SetControltext('MONTANT9',Q.FindField('PFD_E9').AsString);
       SetControltext('MONTANT11',Q.FindField('PFD_F10').AsString);
       SetControltext('MONTANT11A',Q.FindField('PFD_F10A').AsString);
       SetControltext('MONTANT11B',Q.FindField('PFD_F10B').AsString);
       SetControltext('MONTANT11C',Q.FindField('PFD_F10C').AsString);
       SetControltext('MONTANT11D',Q.FindField('PFD_F10D').AsString);
       SetControltext('MONTANT11E',Q.FindField('PFD_F10E').AsString);
       SetControltext('MONTANT11F',Q.FindField('PFD_F10F').AsString);
       SetControltext('MONTANT11G',Q.FindField('PFD_F10G').AsString);
       SetControltext('MONTANT11H',Q.FindField('PFD_F10H').AsString);
       SetControltext('MONTANT11I',Q.FindField('PFD_F11').AsString);
       SetControltext('MONTANT12',Q.FindField('PFD_F12').AsString);
       SetControltext('MONTANT13',Q.FindField('PFD_F13').AsString);
       SetControltext('MONTANT14',Q.FindField('PFD_F14').AsString);
       SetControltext('MONTANT15',Q.FindField('PFD_F15').AsString);
       SetControltext('MONTANT16',Q.FindField('PFD_G16').AsString);
       SetControltext('MONTANT17',Q.FindField('PFD_G17').AsString);
       SetControltext('MONTANT18',Q.FindField('PFD_G18').AsString);
       SetControltext('MONTANT19',Q.FindField('PFD_G19').AsString);
       SetControltext('MONTANT20',Q.FindField('PFD_H20').AsString);
       SetControltext('MONTANT21',Q.FindField('PFD_H21').AsString);
       SetControltext('MONTANT22',Q.FindField('PFD_H22').AsString);
       SetControltext('MONTANT23',Q.FindField('PFD_H23').AsString);
       SetControltext('MONTANT24',Q.FindField('PFD_H24').AsString);
       SetControltext('EXCEDENTA',Q.FindField('PFD_I2A').AsString);
       SetControltext('EXCEDENTB',Q.FindField('PFD_I2B').AsString);
       SetControltext('EXCEDENTC',Q.FindField('PFD_I3C').AsString);
       SetControltext('EXCEDENTD',Q.FindField('PFD_I3D').AsString);
       SetControltext('EXCEDENTE',Q.FindField('PFD_I4E').AsString);
       SetControltext('TOTAL',Q.FindField('PFD_CEDITIMPOT1').AsString);
       SetControltext('HONNEUR',Q.FindField('PFD_ATTESTHON').AsString);
       SetControltext('JUSTIFIER',Q.FindField('PFD_OBLIG').AsString);
  end;
  Ferme(Q);
  //PT6
  CalcExcedents(GetControl('EXCEDENTA'));
  CalcExcedents(GetControl('EXCEDENTB'));
  CalcExcedents(GetControl('EXCEDENTE'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 01/06/2007 / PT5
Modifié le ... :   /  /
Description .. : Affiche le détail du calcul du champ Montant11A
Mots clefs ... :
*****************************************************************}
procedure TOF_PGDECLARAFORMATION.EditDetailsA (Sender : TObject);
Var
     TobEtat,TobDetails,T : TOB;
     Q       : TQuery;
     Total   : Double;
     i       : Integer;
     Exportation : Boolean;
     {$IFDEF EAGL}StPages : String;{$ENDIF}
begin
        TobEtat := Tob.Create('TobEdition',Nil,-1);

        // Salaires des animateurs
        Q := OpenSQL('SELECT PAN_SALARIE, PAN_CODESTAGE, PAN_ORDRE, PAN_SALAIREANIM,PSS_LIBELLE,PSS_DATEDEBUT,PSS_DATEFIN FROM SESSIONANIMAT '+
        'LEFT JOIN SESSIONSTAGE ON PAN_CODESTAGE=PSS_CODESTAGE AND PAN_ORDRE=PSS_ORDRE AND PAN_MILLESIME=PSS_MILLESIME '+
        'WHERE PAN_SALAIREANIM>0 AND PSS_EFFECTUE="X" AND '+
        'PSS_DATEFIN>="'+UsDateTime(DateDeb)+'" AND PSS_DATEFIN<="'+UsDateTime(DateFin)+'" ORDER BY PAN_SALARIE, PAN_CODESTAGE, PAN_ORDRE',True);
        TobDetails := TOB.Create('TobFille', Nil, -1);
        TobDetails.LoadDetailDB('TobFille','','',Q,False);
        Ferme(Q);
        For i := TobDetails.Detail.Count - 1 DownTo 0 Do
        Begin
             T := TOB.Create('TobFille', TobEtat, -1);
             T.AddChampSupValeur('NATURE',        'Salaires des animateurs');
             T.AddChampSupValeur('PSA_SALARIE',   TobDetails.Detail[i].GetValue('PAN_SALARIE'));
             T.AddChampSupValeur('NB_STAGIAIRES', 0);
             T.AddChampSupValeur('PFO_CODESTAGE', TobDetails.Detail[i].GetValue('PAN_CODESTAGE'));
             T.AddChampSupValeur('PFO_ORDRE',     TobDetails.Detail[i].GetValue('PAN_ORDRE'));
             T.AddChampSupValeur('TYPE',          '');
             T.AddChampSupValeur('COUT',          TobDetails.Detail[i].GetValue('PAN_SALAIREANIM'));
             T.AddChampSupValeur('PSS_LIBELLE',   TobDetails.Detail[i].GetValue('PSS_LIBELLE'));
             T.AddChampSupValeur('PSS_DATEDEBUT', TobDetails.Detail[i].GetValue('PSS_DATEDEBUT'));
             T.AddChampSupValeur('PSS_DATEFIN',   TobDetails.Detail[i].GetValue('PSS_DATEFIN'));
        End;
        FreeAndNil(TobDetails);

        // Frais des stagiaires et animateurs
        Q := OpenSQL('SELECT PFS_SALARIE, PFS_CODESTAGE,PFS_ORDRE, PFS_FRAISSALFOR, PFS_MONTANT, PSS_LIBELLE,PSS_DATEDEBUT,PSS_DATEFIN FROM FRAISSALFORM '+
        'LEFT JOIN SESSIONSTAGE ON PFS_CODESTAGE=PSS_CODESTAGE AND PFS_ORDRE=PSS_ORDRE AND PFS_MILLESIME=PSS_MILLESIME '+
        'WHERE PFS_MONTANT>0 AND PSS_EFFECTUE="X" AND '+
        'PSS_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PSS_DATEDEBUT<="'+UsDateTime(DateFin)+'" ' +
        'ORDER BY PFS_SALARIE,PFS_CODESTAGE,PFS_ORDRE,PFS_FRAISSALFOR',True);
        TobDetails := TOB.Create('TobFille', Nil, -1);
        TobDetails.LoadDetailDB('TobFille','','',Q,False);
        Ferme(Q);
        For i := TobDetails.Detail.Count - 1 DownTo 0 Do
        Begin
             T := TOB.Create('TobFille', TobEtat, -1);
             T.AddChampSupValeur('NATURE',        'Frais des stagiaires et animateurs');
             T.AddChampSupValeur('PSA_SALARIE',   TobDetails.Detail[i].GetValue('PFS_SALARIE'));
             T.AddChampSupValeur('NB_STAGIAIRES', 0);
             T.AddChampSupValeur('PFO_CODESTAGE', TobDetails.Detail[i].GetValue('PFS_CODESTAGE'));
             T.AddChampSupValeur('PFO_ORDRE',     TobDetails.Detail[i].GetValue('PFS_ORDRE'));
             T.AddChampSupValeur('TYPE',          RechDom('PGFRAISSALFORM',TobDetails.Detail[i].GetValue('PFS_FRAISSALFOR'),False));
             T.AddChampSupValeur('COUT',          TobDetails.Detail[i].GetValue('PFS_MONTANT'));
             T.AddChampSupValeur('PSS_LIBELLE',   TobDetails.Detail[i].GetValue('PSS_LIBELLE'));
             T.AddChampSupValeur('PSS_DATEDEBUT', TobDetails.Detail[i].GetValue('PSS_DATEDEBUT'));
             T.AddChampSupValeur('PSS_DATEFIN',   TobDetails.Detail[i].GetValue('PSS_DATEFIN'));
        End;
        FreeAndNil(TobDetails);

        // Cout pédagogique
        Q := OpenSQL('SELECT COUNT(PFO_SALARIE) AS NBSTA, PSS_CODESTAGE,PSS_ORDRE,PSS_TYPCONVFORM,PSS_COUTPEDAG,PSS_COUTUNITAIRE,PSS_LIBELLE,PSS_DATEDEBUT,PSS_DATEFIN '+
        'FROM SESSIONSTAGE LEFT JOIN FORMATIONS ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE AND PSS_MILLESIME=PFO_MILLESIME '+
        'WHERE PSS_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PSS_DATEDEBUT<="'+UsDateTime(DateFin)+
        '" AND PSS_EFFECTUE="X" AND PFO_EFFECTUE="X" AND PSS_INCLUSDECL="X" AND PSS_NATUREFORM="001" '+
        'AND (PSS_COUTUNITAIRE>0 OR PSS_COUTPEDAG>0) '+
        'GROUP BY PSS_TYPCONVFORM,PSS_NATUREFORM,PSS_COUTMOB,PSS_COUTFONCT,PSS_COUTORGAN,PSS_COUTPEDAG,PSS_COUTUNITAIRE,PSS_CODESTAGE,PSS_ORDRE,PSS_LIBELLE,PSS_DATEDEBUT,PSS_DATEFIN',True); //PT1 et PT2
        TobDetails := Tob.Create('TobFille',Nil,-1);
        TobDetails.LoadDetailDB('TobFille','','',Q,False);
        Ferme(Q);
        For i := TobDetails.Detail.Count - 1 DownTo 0 Do
        begin
             T := TOB.Create('TobFille', TobEtat, -1);
             T.AddChampSupValeur('NATURE',        'Coût pédagogique');
             T.AddChampSupValeur('PSA_SALARIE',   '');
             T.AddChampSupValeur('NB_STAGIAIRES', IntToStr(TobDetails.Detail[i].GetValue('NBSTA')));  // pour affichage dans l'état, il faut du string
             T.AddChampSupValeur('PFO_CODESTAGE', TobDetails.Detail[i].GetValue('PSS_CODESTAGE'));
             T.AddChampSupValeur('PFO_ORDRE',     TobDetails.Detail[i].GetValue('PSS_ORDRE'));
             T.AddChampSupValeur('TYPE',          RechDom('PGTYPCONVFORM',TobDetails.Detail[i].GetValue('PSS_TYPCONVFORM'),False));
             Total := TobDetails.Detail[i].GetValue('PSS_COUTPEDAG') +
                    (TobDetails.Detail[i].GetValue('PSS_COUTUNITAIRE') * TobDetails.detail[i].GetValue('NBSTA'));
             T.AddChampSupValeur('COUT',          Total);
             T.AddChampSupValeur('PSS_LIBELLE',   TobDetails.Detail[i].GetValue('PSS_LIBELLE'));
             T.AddChampSupValeur('PSS_DATEDEBUT', TobDetails.Detail[i].GetValue('PSS_DATEDEBUT'));
             T.AddChampSupValeur('PSS_DATEFIN',   TobDetails.Detail[i].GetValue('PSS_DATEFIN'));
        end;
        FreeAndNil(TobDetails);

        // Intitulé de l'état
        SetControlText('LIBELLE_COUT','les formations internes');

        // Lancement de l'état
        If (TMenuItem(Sender).Name = 'EDITIONA') Then Exportation := False Else Exportation := True;
        SetControlChecked('FLISTE',Exportation);
        {$IFDEF EAGL}
        StPages := AGLGetCriteres(TPageControl(GetControl('PAGE1_')));
        LanceEtatTOB('E','PFO','PDF',TobEtat,True,Exportation,False,TPageControl(GetControl('PAGE1_')),'','',False,0,StPages);
        {$ELSE}
        LanceEtatTOB('E','PFO','PDF',TobEtat,True,Exportation,False,TPageControl(GetControl('PAGE1_')),'','',False,0,'');
        {$ENDIF}

        FreeAndNil(TobEtat);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 01/06/2007 / PT5
Modifié le ... :   /  /
Description .. : Affiche le détail du calcul du champ Montant11B
Mots clefs ... :
*****************************************************************}
procedure TOF_PGDECLARAFORMATION.EditDetailsB (Sender : TObject);
Var
     TobEtat,TobDetails,T : TOB;
     Q       : TQuery;
     Total   : Double;
     i       : Integer;
     Exportation : Boolean;
     {$IFDEF EAGL}StPages : String;{$ENDIF}
begin
        //If (TMenuItem(Sender).Name = 'EDITIONB') Then Exportation := False
        //Else Exportation := True;

        TobEtat := Tob.Create('TobEdition',Nil,-1);

        // Cout pédagogique
        Q := OpenSQL('SELECT COUNT(PFO_SALARIE) AS NBSTA, PSS_CODESTAGE,PSS_ORDRE,PSS_TYPCONVFORM,PSS_COUTPEDAG,PSS_COUTUNITAIRE,PSS_LIBELLE,PSS_DATEDEBUT,PSS_DATEFIN '+
        'FROM SESSIONSTAGE LEFT JOIN FORMATIONS ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE AND PSS_MILLESIME=PFO_MILLESIME '+
        'WHERE PSS_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND PSS_DATEDEBUT<="'+UsDateTime(DateFin)+
        '" AND PSS_EFFECTUE="X" AND PFO_EFFECTUE="X" AND PSS_INCLUSDECL="X" AND PSS_NATUREFORM<>"001" '+
        'AND (PSS_COUTUNITAIRE>0 OR PSS_COUTPEDAG>0) '+
        'GROUP BY PSS_TYPCONVFORM,PSS_NATUREFORM,PSS_COUTMOB,PSS_COUTFONCT,PSS_COUTORGAN,PSS_COUTPEDAG,PSS_COUTUNITAIRE,PSS_CODESTAGE,PSS_ORDRE,PSS_LIBELLE,PSS_DATEDEBUT,PSS_DATEFIN',True); //PT1 et PT2
        TobDetails := Tob.Create('TobFille',Nil,-1);
        TobDetails.LoadDetailDB('TobFille','','',Q,False);
        Ferme(Q);
        For i := TobDetails.Detail.Count - 1 DownTo 0 Do
        begin
             T := TOB.Create('TobFille', TobEtat, -1);
             T.AddChampSupValeur('NATURE',        'Coût pédagogique');
             T.AddChampSupValeur('PSA_SALARIE',   '');
             T.AddChampSupValeur('NB_STAGIAIRES', IntToStr(TobDetails.Detail[i].GetValue('NBSTA')));  // pour affichage dans l'état, il faut du string
             T.AddChampSupValeur('PFO_CODESTAGE', TobDetails.Detail[i].GetValue('PSS_CODESTAGE'));
             T.AddChampSupValeur('PFO_ORDRE',     TobDetails.Detail[i].GetValue('PSS_ORDRE'));
             T.AddChampSupValeur('TYPE',          RechDom('PGTYPCONVFORM',TobDetails.Detail[i].GetValue('PSS_TYPCONVFORM'),False));
             Total := TobDetails.Detail[i].GetValue('PSS_COUTPEDAG') +
                    (TobDetails.Detail[i].GetValue('PSS_COUTUNITAIRE') * TobDetails.detail[i].GetValue('NBSTA'));
             T.AddChampSupValeur('COUT',          Total);
             T.AddChampSupValeur('PSS_LIBELLE',   TobDetails.Detail[i].GetValue('PSS_LIBELLE'));
             T.AddChampSupValeur('PSS_DATEDEBUT', TobDetails.Detail[i].GetValue('PSS_DATEDEBUT'));
             T.AddChampSupValeur('PSS_DATEFIN',   TobDetails.Detail[i].GetValue('PSS_DATEFIN'));
        end;
        FreeAndNil(TobDetails);

        // Intitulé de l'état
        SetControlText('LIBELLE_COUT','les formations externes');

        // Lancement de l'état
        If (TMenuItem(Sender).Name = 'EDITIONB') Then Exportation := False Else Exportation := True;
        SetControlChecked('FLISTE',Exportation);
        {$IFDEF EAGL}
        StPages := AGLGetCriteres(TPageControl(GetControl('PAGE1_')));
        LanceEtatTOB('E','PFO','PDF',TobEtat,True,Exportation,False,TPageControl(GetControl('PAGE1_')),'','',False,0,StPages);
        {$ELSE}
        LanceEtatTOB('E','PFO','PDF',TobEtat,True,Exportation,False,TPageControl(GetControl('PAGE1_')),'','',False,0,'');
        {$ENDIF}

        FreeAndNil(TobEtat);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 01/06/2007 / PT5
Modifié le ... :   /  /
Description .. : Affiche le détail du calcul du champ Montant11D
Mots clefs ... :
*****************************************************************}
procedure TOF_PGDECLARAFORMATION.EditDetailsD (Sender : TObject);
Var
     TobEtat,TobDetails,T : TOB;
     Q       : TQuery;
//     Total   : Double;
     i       : Integer;
     Exportation : Boolean;
     {$IFDEF EAGL}StPages : String;{$ENDIF}
begin
        //If (TMenuItem(Sender).Name = 'EDITIOND') Then Exportation := False
        //Else Exportation := True;

        TobEtat := Tob.Create('TobEdition',Nil,-1);

        // Allocations
        Q := OpenSQL('SELECT PFO_CODESTAGE, PFO_ORDRE, PFO_SALARIE, PFO_ALLOCFORM, PSS_LIBELLE,PSS_DATEDEBUT,PSS_DATEFIN FROM FORMATIONS'+
        ' LEFT JOIN SESSIONSTAGE ON PFO_CODESTAGE=PSS_CODESTAGE AND PFO_ORDRE=PSS_ORDRE AND PFO_MILLESIME=PSS_MILLESIME '+
        ' WHERE PFO_DATEFIN>="'+UsDateTime(DateDeb)+'" AND PFO_DATEFIN<="'+UsDateTime(DateFin)+'" AND PFO_EFFECTUE="X" AND PFO_ALLOCFORM>0'+
        ' ORDER BY PFO_CODESTAGE, PFO_ORDRE, PFO_SALARIE',True); //PT1 et PT2
        TobDetails := Tob.Create('TobFille',Nil,-1);
        TobDetails.LoadDetailDB('TobFille','','',Q,False);
        Ferme(Q);
        For i := TobDetails.Detail.Count - 1 DownTo 0 Do
        begin
             T := TOB.Create('TobFille', TobEtat, -1);
             T.AddChampSupValeur('NATURE',        'Allocations');
             T.AddChampSupValeur('PSA_SALARIE',   TobDetails.Detail[i].GetValue('PFO_SALARIE'));
             T.AddChampSupValeur('NB_STAGIAIRES', 0);
             T.AddChampSupValeur('PFO_CODESTAGE', TobDetails.Detail[i].GetValue('PFO_CODESTAGE'));
             T.AddChampSupValeur('PFO_ORDRE',     TobDetails.Detail[i].GetValue('PFO_ORDRE'));
             T.AddChampSupValeur('TYPE',          '');
             T.AddChampSupValeur('COUT',          TobDetails.Detail[i].GetValue('PFO_ALLOCFORM'));
             T.AddChampSupValeur('PSS_LIBELLE',   TobDetails.Detail[i].GetValue('PSS_LIBELLE'));
             T.AddChampSupValeur('PSS_DATEDEBUT', TobDetails.Detail[i].GetValue('PSS_DATEDEBUT'));
             T.AddChampSupValeur('PSS_DATEFIN',   TobDetails.Detail[i].GetValue('PSS_DATEFIN'));
        end;
        FreeAndNil(TobDetails);

        // Intitulé de l'état
        SetControlText('LIBELLE_COUT','les allocations');

        // Lancement de l'état
        If (TMenuItem(Sender).Name = 'EDITIOND') Then Exportation := False Else Exportation := True;
        SetControlChecked('FLISTE',Exportation);
        {$IFDEF EAGL}
        StPages := AGLGetCriteres(TPageControl(GetControl('PAGE1_')));
        LanceEtatTOB('E','PFO','PDF',TobEtat,True,Exportation,False,TPageControl(GetControl('PAGE1_')),'','',False,0,StPages);
        {$ELSE}
        LanceEtatTOB('E','PFO','PDF',TobEtat,True,Exportation,False,TPageControl(GetControl('PAGE1_')),'','',False,0,'');
        {$ENDIF}

        FreeAndNil(TobEtat);
end;

Initialization
registerclasses([TOF_PGDECLARAFORMATION]);
end.
