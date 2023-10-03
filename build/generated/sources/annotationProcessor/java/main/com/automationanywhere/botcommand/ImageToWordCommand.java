package com.automationanywhere.botcommand;

import com.automationanywhere.bot.service.GlobalSessionContext;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import java.lang.ClassCastException;
import java.lang.Deprecated;
import java.lang.Object;
import java.lang.String;
import java.lang.Throwable;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public final class ImageToWordCommand implements BotCommand {
  private static final Logger logger = LogManager.getLogger(ImageToWordCommand.class);

  private static final Messages MESSAGES_GENERIC = MessagesFactory.getMessages("com.automationanywhere.commandsdk.generic.messages");

  @Deprecated
  public Optional<Value> execute(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return execute(null, parameters, sessionMap);
  }

  public Optional<Value> execute(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    logger.traceEntry(() -> parameters != null ? parameters.entrySet().stream().filter(en -> !Arrays.asList( new String[] {}).contains(en.getKey()) && en.getValue() != null).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)).toString() : null, ()-> sessionMap != null ?sessionMap.toString() : null);
    ImageToWord command = new ImageToWord();
    HashMap<String, Object> convertedParameters = new HashMap<String, Object>();
    if(parameters.containsKey("sourceFolderPath") && parameters.get("sourceFolderPath") != null && parameters.get("sourceFolderPath").get() != null) {
      convertedParameters.put("sourceFolderPath", parameters.get("sourceFolderPath").get());
      if(convertedParameters.get("sourceFolderPath") !=null && !(convertedParameters.get("sourceFolderPath") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","sourceFolderPath", "String", parameters.get("sourceFolderPath").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("sourceFolderPath") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","sourceFolderPath"));
    }
    if(convertedParameters.containsKey("sourceFolderPath")) {
      String filePath= ((String)convertedParameters.get("sourceFolderPath"));
      int lastIndxDot = filePath.lastIndexOf(".");
      if (lastIndxDot == -1 || lastIndxDot >= filePath.length()) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.FileExtension","sourceFolderPath","jpeg,jpg,gif,png,webp"));
      }
      String fileExtension = filePath.substring(lastIndxDot + 1);
      if(!Arrays.stream("jpeg,jpg,gif,png,webp".split(",")).anyMatch(fileExtension::equalsIgnoreCase))  {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.FileExtension","sourceFolderPath","jpeg,jpg,gif,png,webp"));
      }

    }
    if(parameters.containsKey("destinationFolderPath") && parameters.get("destinationFolderPath") != null && parameters.get("destinationFolderPath").get() != null) {
      convertedParameters.put("destinationFolderPath", parameters.get("destinationFolderPath").get());
      if(convertedParameters.get("destinationFolderPath") !=null && !(convertedParameters.get("destinationFolderPath") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","destinationFolderPath", "String", parameters.get("destinationFolderPath").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("destinationFolderPath") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","destinationFolderPath"));
    }

    try {
      Optional<Value> result =  Optional.ofNullable(command.action((String)convertedParameters.get("sourceFolderPath"),(String)convertedParameters.get("destinationFolderPath")));
      return logger.traceExit(result);
    }
    catch (ClassCastException e) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.IllegalParameters","action"));
    }
    catch (BotCommandException e) {
      logger.fatal(e.getMessage(),e);
      throw e;
    }
    catch (Throwable e) {
      logger.fatal(e.getMessage(),e);
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.NotBotCommandException",e.getMessage()),e);
    }
  }
}
