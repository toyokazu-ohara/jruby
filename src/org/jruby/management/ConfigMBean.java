package org.jruby.management;

import java.util.List;
import java.util.Map;

public interface ConfigMBean {
    public String getVersionString();
    public String getCopyrightString();
    public String getCompileMode();
    public boolean isJitLogging();
    public boolean isJitLoggingVerbose();
    public int getJitLogEvery();
    public boolean isSamplingEnabled();
    public int getJitThreshold();
    public int getJitMax();
    public boolean isRunRubyInProcess();
    public String getCompatVersion();
    public String getCurrentDirectory();
    public boolean isObjectSpaceEnabled();
    public String getEnvironment();
    public String getArgv();
    public String getJRubyHome();
    public String getRequiredLibraries();
    public String getLoadPaths();
    public String getDisplayedFileName();
    public String getScriptFileName();
    public boolean isBenchmarking();
    public boolean isAssumeLoop();
    public boolean isAssumePrinting();
    public boolean isProcessLineEnds();
    public boolean isSplit();
    public boolean isVerbose();
    public boolean isDebug();
    public boolean isYARVEnabled();
    public String getInputFieldSeparator();
    public boolean isRubiniusEnabled();
    public boolean isYARVCompileEnabled();
    public String getKCode();
    public String getRecordSeparator();
    public int getSafeLevel();
    public String getOptionGlobals();
    public boolean isManagementEnabled();
}