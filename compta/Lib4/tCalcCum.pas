unit tcalccum ;

interface

Uses Ent1   ,UentCommun;

Type TTypeCalc=(Un,Deux,DeuxSansCumul,UnSurTableLibre,DeuxSurTableLibre,CPCpteJal,
                CPUn,CPUnUDF,UnBud,DeuxBud,UnNatEcr,DeuxNatEcr) ;

type ZSParams = record
     fb1,
     fb2        : TFichierBase ;
     SetTyp     : SetttTypePiece ;
     Multiple   : TTypeCalc ;
     SurQuelCpI,
     SurQuelCpO,
     Exo,
     Etab,
     Devise,
     TypeCumul,
     TypePlan   : string ;
     Date1,
     Date2      : TDateTime ;
     DevPivot,
     bSQLCumul  : boolean;
     end ;

type PZSParams = ^ZSParams ;

Type TabDC = RECORD
             TotDebit  : Double ;
             TotCredit : Double ;
             END ;

Type TabTot = Array[0..7] Of TabDC ;



implementation

end.


