/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package org.jruby.ext.win32ole;


import org.jruby.*;
import org.jruby.anno.JRubyClass;
import org.jruby.anno.JRubyMethod;
import org.jruby.runtime.Block;
import org.jruby.runtime.ObjectAllocator;
import org.jruby.runtime.ThreadContext;
import org.jruby.runtime.Visibility;
import org.jruby.runtime.builtin.IRubyObject;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import org.jruby.runtime.load.Library;




/**
 * @author toyokazu.ohara
 */
@JRubyClass(name="WIN32OLE")
public class RubyWIN32OLE extends RubyObject {
    private static final int INVOKE_FUNC = ActiveXComponent.Method;
    private static final int INVOKE_PROPERTYGET = ActiveXComponent.Get;
    private static final int INVOKE_PROPERTYPUT = ActiveXComponent.Put;
    private static final int INVOKE_PROPERTYPUTREF = ActiveXComponent.Put;
    private ActiveXComponent ax = null;
    private Dispatch  dispatch = null;

    public static RubyClass createWIN32OLE(Ruby runtime) {
        RubyClass result = runtime.defineClass("WIN32OLE", runtime.getObject(), WIN32OLE_ALLOCATOR);
        result.kindOf = new RubyModule.KindOf() {
            @Override
            public boolean isKindOf(IRubyObject obj, RubyModule type) {
                return obj instanceof RubyWIN32OLE;
            }
        };

        result.defineAnnotatedMethods(RubyWIN32OLE.class);
        return result;
    }

    private static final ObjectAllocator WIN32OLE_ALLOCATOR = new ObjectAllocator() {
        public IRubyObject allocate(Ruby runtime, RubyClass klass) {
            return new RubyWIN32OLE(runtime, klass);
        }
    };

    private RubyWIN32OLE(Ruby runtime, RubyClass klass) {
        super(runtime, klass);
    }

    @JRubyMethod(name = "initialize", required = 1, optional = 1, frame = true, visibility = Visibility.PRIVATE)
    public IRubyObject initialize(IRubyObject[] args, Block unusedBlock) {
        RubyString server = args[0].convertToString();
        ax = new ActiveXComponent(server.asJavaString());
        return this;
    }

// ----- Ruby Class Methods ----------------------------------------------------

    @JRubyMethod(name = "method_missing", rest = true, frame = true, module = true, visibility = Visibility.PRIVATE)
    public static IRubyObject method_missing(ThreadContext context, IRubyObject recv, IRubyObject[] args, Block block) {
try {
        if (args.length == 0 || !(args[0] instanceof RubySymbol)) throw context.getRuntime().newArgumentError("no id given");
        System.out.println("length : " + args.length);
        for (int i = 0; i < args.length; i++) {
            System.out.println("arg [" + i + "] : value = " + args[i] + " type : " + args[i].getClass());
        }

        RubyWIN32OLE self  = (RubyWIN32OLE) recv;
        String methodName  = args[0].asJavaString();
        Ruby runtime       = context.getRuntime();


        if (methodName.endsWith("=")) {
            String realName = methodName.substring(0, methodName.length() - 1);
            if (self.dispatch == null) {
                setProperty(self.ax.getObject(), realName, args[1]);
            } else {
                setProperty(self.dispatch, realName, args[1]);
            }
            return args[1];
        } else {
            Dispatch dispatch;
            if (self.dispatch == null) {
                dispatch = self.ax.getObject();
            } else {
                dispatch = self.dispatch;
            }
            Object[] win32args = getWin32Args(dispatch, args, true);
            Variant result = Dispatch.callN(dispatch, methodName, win32args);
            return getVariantValue(runtime, result, self);
        }

} catch (Exception e) {
    e.printStackTrace();
    return null;
}
    }

    @JRubyMethod(name = "[]=", frame = true, visibility = Visibility.PUBLIC)
    public IRubyObject aset(IRubyObject arg0, IRubyObject arg1) {
        if (dispatch == null) {
            setProperty(ax.getObject(), arg0.asJavaString(), arg1);
        } else {
            setProperty(dispatch, arg0.asJavaString(), arg1);
        }
        return arg1;
    }


//
//    @JRubyMethod(name = "ole_method_help", frame = true, visibility = Visibility.PUBLIC)
//    public IRubyObject ole_method_help(IRubyObject[] args, Block unusedBlock) {
//        int dispID = INVOKE_PROPERTYGET;
//        return invokeMethod(args, unusedBlock, dispID);
//    }
//
//
//    @JRubyMethod(name = "ole_obj_help", frame = true, visibility = Visibility.PUBLIC)
//    public IRubyObject ole_obj_help(IRubyObject[] args, Block unusedBlock) {
//        int dispID = INVOKE_PROPERTYGET;
//        return invokeMethod(args, unusedBlock, dispID);
//    }
//
//
//
//    @JRubyMethod(name = "ole_methods", frame = true, visibility = Visibility.PUBLIC)
//    public IRubyObject ole_methods(IRubyObject[] args, Block unusedBlock) {
//        int dispID = INVOKE_FUNC | INVOKE_PROPERTYGET | INVOKE_PROPERTYPUT | INVOKE_PROPERTYPUTREF;
//        return invokeMethod(args, unusedBlock, dispID);
//    }

    public IRubyObject invokeMethod(IRubyObject[] args, Block unusedBlock, int dispID) {
        Dispatch invokeTarget;
        if (this.dispatch == null) {
            invokeTarget = ax.getObject();
        } else {
            invokeTarget = this.dispatch;
        }

        Object[] win32args = getWin32Args(invokeTarget, args, false);
        Variant result = Dispatch.callN(invokeTarget, dispID, win32args);
        return getVariantValue(getRuntime(), result, this);
    }




    private static void setProperty(Dispatch dispatch, String propName, IRubyObject value) {
        if (value instanceof RubyMethod) {
            // TODO for method
        } else {
            Dispatch.put(dispatch, propName, getWin32Arg(value));
        }
    }


    private static Object[] getWin32Args(Dispatch dispatch, IRubyObject[] args, boolean isIgnoreFirst) {
        int skip;
        int length;
        if (isIgnoreFirst) {
            skip   = 1;
            length = args.length - 1;
        } else {
            skip   = 0;
            length = args.length;
        }

        List result = new ArrayList();
        for (int i = 0; i < length; i++) {
            IRubyObject arg = args[i + skip];
            if (arg instanceof RubyHash) {
                result.addAll(Arrays.asList(getWin32Arg(dispatch, (RubyHash) arg)));
            } else {
                result.add(getWin32Arg(arg));
            }
        }
        return result.toArray();
    }

    private static Object getWin32Arg(IRubyObject arg) {
        if (arg instanceof RubyWIN32OLE) {
            RubyWIN32OLE real = (RubyWIN32OLE) arg;
            if (real.dispatch == null) {
                return real.ax;
            } else {
                return real.dispatch;
            }
        } else if (arg instanceof RubyFixnum) {
            RubyFixnum real = (RubyFixnum) arg;
            return Long.valueOf(real.getLongValue());
        } else if (arg instanceof RubyFloat) {
            RubyFloat real = (RubyFloat) arg;
            return Double.valueOf(real.getDoubleValue());
        } else if (arg instanceof RubyBignum) {
            RubyBignum real = (RubyBignum) arg;
            return real.getValue();
        } else if (arg instanceof RubyBigDecimal) {
            RubyBigDecimal real = (RubyBigDecimal) arg;
            return real.getValue();
        } else if (arg instanceof RubyBoolean) {
            RubyBoolean real = (RubyBoolean) arg;
            return Boolean.valueOf(arg.isTrue());
        } else if (arg instanceof RubyString) {
            return arg.asJavaString();
        } else if (arg instanceof RubyString) {
            return arg.asJavaString();
        } else if (arg instanceof RubyArray) {
            RubyArray real  = (RubyArray) arg;
            Object[] result = new Object[real.getLength()];
            for (int i = 0, iMax = real.getLength(); i < iMax; i++) {
                result[i] = getWin32Arg(real.entry(i));
            }
            return result;
        } else {
            return arg;
        }
    }

    private static Object[] getWin32Arg(Dispatch dispatch, RubyHash arg) {
System.out.println("RubyHash");
        String[] names  = new String[arg.directKeySet().size()];
        Object[] values = new Object[arg.directKeySet().size()];
        int pos         = 0;
        for (Iterator it = arg.directKeySet().iterator(); it.hasNext();) {
            IRubyObject key   = (IRubyObject) it.next();
            IRubyObject value = (IRubyObject) arg.fastARef(key);
            names[pos]        = key.asJavaString();
            values[pos]       = getWin32Arg(value);
System.out.println("key : " + key + " value : " + value);
            pos ++;
        }
        int[] ids  = Dispatch.getIDsOfNames(dispatch, names);
        int length = 0;
        for (int i = 0; i < ids.length; i++) {
            int id = ids[i];
            if (id > length) {
                length = id;
            }
        }
        Object[] result = new Object[length];
        for (int i = 0; i < ids.length; i++) {
            result[ids[i]] = values[i];
        }

        return result;
    }

    private static IRubyObject getVariantValue(Ruby runtime, Variant variant, RubyWIN32OLE self){
        IRubyObject result = null;
        System.out.println("variant type: " + variant.getvt());
        // TODO check bindings
        switch(variant.getvt()) {
            case Variant.VariantEmpty://0
                result =  runtime.getNil();
                break;
            case Variant.VariantNull://1
                result =  runtime.getNil();
                break;
            case Variant.VariantShort://2
                result =  runtime.newFixnum(variant.getShort());
                break;
            case Variant.VariantInt://3
                result =  runtime.newFixnum(variant.getInt());
                break;
            case Variant.VariantFloat://4
                result =  runtime.newFloat(variant.getFloat());
                break;
            case Variant.VariantDouble://5
                result =  runtime.newFloat(variant.getDouble());
                break;
            case Variant.VariantCurrency://6
                // TODO
                break;
            case Variant.VariantDate://7
                // TODO
                break;
            case Variant.VariantString://8
                result =  runtime.newString(variant.getString());
                break;
            case Variant.VariantDispatch://9
                RubyWIN32OLE rubyObject = new RubyWIN32OLE(runtime, createWIN32OLE(runtime));
                rubyObject.ax        = self.ax;
                rubyObject.dispatch = variant.toDispatch();
                result =  rubyObject;
                break;
            case Variant.VariantError://10
                // TODO
                break;
            case Variant.VariantBoolean://11
                result =  runtime.newBoolean(variant.getBoolean());
                break;
            case Variant.VariantVariant://12
                // TODO
                break;
            case Variant.VariantObject://13
                // TODO
                break;
            case Variant.VariantDecimal://14
                result =  new RubyBigDecimal(runtime, variant.getDecimal());
                break;
            case Variant.VariantByte://17
                // TODO
                break;
            case Variant.VariantLongInt://20
                result =  new RubyFixnum(runtime, variant.getLong());
                break;
            default:
            break;
        }
        if (result == null) {
            result =  self;
        }
        System.out.println("result : " + result);
        return result;
    }







    @Override
    protected void finalize() throws Throwable {
        if (dispatch == null) {
            // TODO close
        }
        super.finalize();
    }

    public static class WIN32OLELibrary implements Library {
        public void load(final Ruby runtime, boolean wrap) throws IOException {
            RubyWIN32OLE.createWIN32OLE(runtime);
        }
    }


//    @JRubyClass(name="WIN32OLE_TYPE")
//    public static class RubyWIN32OLE_TYPE extends RubyObject {
//        public static RubyClass createWIN32OLE_TYPE(Ruby runtime) {
//            RubyClass result = runtime.defineClass("WIN32OLE_TYPE", runtime.getObject(), WIN32OLE_TYPE_ALLOCATOR);
//            result.kindOf = new RubyModule.KindOf() {
//                @Override
//                public boolean isKindOf(IRubyObject obj, RubyModule type) {
//                    return obj instanceof RubyWIN32OLE_TYPE;
//                }
//            };
//            result.defineAnnotatedMethods(RubyWIN32OLE_TYPE.class);
//            return result;
//        }
//        private static final ObjectAllocator WIN32OLE_TYPE_ALLOCATOR = new ObjectAllocator() {
//            public IRubyObject allocate(Ruby runtime, RubyClass klass) {
//                return new RubyWIN32OLE_TYPE(runtime, klass);
//            }
//        };
//        private RubyWIN32OLE_TYPE(Ruby runtime, RubyClass klass) {
//            super(runtime, klass);
//        }
//    }
}
