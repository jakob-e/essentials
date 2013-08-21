/*
    Financial functions based on formula.js
    http://stoic.com/formula/lib/formula.js

        PV(rate,nper,pmt,fv,type)
        FV(rate,nper,pmt,pv,type)
        RATE(nper,pmt,pv,fv,type,guess)
        PMT(rate,nper,pv,fv,type)
        PPMT(rate,per,nper,pv,fv,type)
        IPMT(rate,per,nper,pv,fv,type)
        NPER(rate,pmt,pv,fv,type)
        EFFECT(nominal_rate,npery)
        NOMINAL(effect_rate,npery)
*/
var excel={
    // Present Value
    // -----------------------------
    // Returns the present value of an investment. The present value is the 
    // total amount that a series of future payments is worth now. For example,
    // when you borrow money,the loan amount is the present value to the lender.
    // 
    // Syntax
    // PV(rate,nper,pmt,fv,type)
    PV:function(rate,periods,payment,future,type){
        // Initialize type
        var type=(typeof type==='undefined')?0:type;
        // Return present value
        if(rate===0){
            return -payment*periods-future;
        }else{
            return (((1-Math.pow(1+rate,periods))/rate)*payment*(1+rate*type)-future)/Math.pow(1+rate,periods);
        }
    }, 
    // Future Value
    // -----------------------------
    // Returns the future value of an investment based on periodic,constant 
    // payments and a constant interest rate.
    // 
    // Syntax
    // FV(rate,nper,pmt,pv,type)
    FV:function(rate,periods,payment,value,type){
        // Initialize type
        var type=(typeof type==='undefined')?0:type;
        // Return future value
        var result;
        if(rate===0){
            result=value+payment*periods;
        }else{
            var term=Math.pow(1+rate,periods);
            if(type===1){
                result=value*term+payment*(1+rate)*(term-1.0)/rate;
            }else{
                result=value*term+payment*(term-1)/rate;
            }
        }
        return -result;
    },
    // Interest rate
    // -----------------------------
    // Returns the interest rate per period of an annuity. RATE is calculated by 
    // iteration and can have zero or more solutions. If the successive results 
    // of RATE do not converge to within 0.0000001 after 20 iterations,RATE 
    // returns the #NUM! error value.
    // 
    // Syntax
    // RATE(nper,pmt,pv,fv,type,guess)
    RATE:function(periods,payment,present,future,type,guess){
        // Initialize guess
        var guess=(typeof guess==='undefined')?0.01:guess;
        // Initialize future
        var future=(typeof future==='undefined')?0:future;
        // Initialize type
        var type=(typeof type==='undefined')?0:type;
        // Set maximum epsilon for end of iteration
        var epsMax=1e-10;
        // Set maximum number of iterations
        var iterMax=50;
        // Implement Newton's method
        var y,y0,y1,x0,x1=0,
            f=0,
            i=0;
        var rate=guess;
        if(Math.abs(rate)<epsMax){
            y=present*(1+periods*rate)+payment*(1+rate*type)*periods+future;
        }else{
            f=Math.exp(periods*Math.log(1+rate));
            y=present*f+payment*(1/rate+type)*(f-1)+future;
        }
        y0=present+payment*periods+future;
        y1=present*f+payment*(1/rate+type)*(f-1)+future;
        i=x0=0;
        x1=rate;
        while ((Math.abs(y0-y1) > epsMax) && (i<iterMax)){
            rate=(y1*x0-y0*x1)/(y1-y0);
            x0=x1;
            x1=rate;
            if(Math.abs(rate)<epsMax){
                y=present*(1+periods*rate)+payment*(1+rate*type)*periods+future;
            }else{
                f=Math.exp(periods*Math.log(1+rate));
                y=present*f+payment*(1/rate+type)*(f-1)+future;
            }
            y0=y1;
            y1=y;
            ++i;
        }
        return rate;
    },
    // Loan Payment
    // -----------------------------
    // Calculates the payment for a loan based on constant payments and 
    // a constant interest rate.
    // 
    // Syntax
    // PMT(rate,nper,pv,fv,type)
    PMT:function(rate,periods,present,future,type){
        // Initialize type
        var type=(typeof type==='undefined')?0:type;
        var future=(typeof future==='undefined')?0:future;
        // Return payment
        var result;
        if(rate===0){
            result=(present+future)/periods;
        }else{
            var term=Math.pow(1+rate,periods);
            if(type===1){
                result=(future*rate/(term-1)+present*rate/(1-1/term))/(1+rate);
            }else{
                result=future*rate/(term-1)+present*rate/(1-1/term);
            }
        }
        return -result;
    },
    // Principal Payment (period)
    // -----------------------------
    // Returns the payment on the principal for a given period for an investment 
    // based on periodic,constant payments and a constant interest rate.
    // 
    // Syntax
    // PPMT(rate,per,nper,pv,fv,type)
    PPMT:function(rate,period,periods,present,future,type){
        return excel.PMT(rate,periods,present,future,type)-excel.IPMT(rate,period,periods,present,future,type);
    },
    // Interest Payment (period)
    // -----------------------------
    // Returns the interest payment for a given period for an investment 
    // based on periodic,constant payments and a constant interest rate.
    // 
    // Syntax
    // IPMT(rate,per,nper,pv,fv,type)
    IPMT:function(rate,period,periods,present,future,type){
        // Initialize type
        var type=(typeof type==='undefined')?0:type;
        // Compute payment
        var payment=excel.PMT(rate,periods,present,future,type);
        // Compute interest
        var interest;
        if(period===1){
            if(type===1){
                interest=0;
            }else{
                interest=-present;
            }
        }else{
            if(type===1){
                interest=excel.FV(rate,period-2,payment,present,1)-payment;
            }else{
                interest=excel.FV(rate,period-1,payment,present,0);
            }
        }
        // Return interest
        return interest*rate;
    },
    // Number of Payments
    // -----------------------------
    // Returns the number of periods for an investment based on periodic,
    // constant payments and a constant interest rate.
    // 
    // Syntax
    // NPER(rate,pmt,pv,fv,type)
    NPER:function(rate,payment,present,future,type){
        // Initialize type
        var type=(typeof type==='undefined')?0:type;
        // Initialize future value
        var future=(typeof future==='undefined')?0:future;
        // Return number of periods
        var num=payment*(1+rate*type)-future*rate;
        var den=(present*rate+payment*(1+rate*type));
        return Math.log(num/den)/Math.log(1+rate);
    },
    // Effective annual interest rate
    // -------------------------------
    // Returns the effective annual interest rate, given the nominal annual interest rate 
    // and the number of compounding periods per year.
    // 
    // Syntax
    // EFFECT(nominal_rate,npery)
    EFFECT:function(rate,periods) {
        // Return error if any of the parameters is not a number
        if(isNaN(rate)||isNaN(periods)) return '#VALUE!';
        // Return error if rate <=0 or periods<1
        if(rate<= 0||periods<1) return '#NUM!';
        // Truncate periods if it is not an integer
        periods=parseInt(periods,10);
        // Return effective annual interest rate
        return Math.pow(1+rate/periods,periods)-1;
    },
    // Nominal annual interest rate
    // -------------------------------
    // Returns the nominal annual interest rate, given the effective rate and the number 
    // of compounding periods per year.
    // 
    // Syntax
    // NOMINAL(effect_rate,npery)
    NOMINAL:function (rate, periods) {
        // Return error if any of the parameters is not a number
        if(isNaN(rate)||isNaN(periods)) return '#VALUE!';
        // Return error if rate <=0 or periods<1
        if(rate<=0||periods<1) return '#NUM!';
        // Truncate periods if it is not an integer
        periods=parseInt(periods,10);
        // Return nominal annual interest rate
        return (Math.pow(rate+1,1/periods)-1)*periods;
    }
}