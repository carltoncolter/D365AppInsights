﻿using System;
using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using System.Collections.Generic;

namespace AppInsightsAction.Tests
{
    [TestClass]
    public class ActionEventTests
    {
        [TestMethod]
        public void ActionEventTest()
        {
            AiSecureConfig aiSecureConfig =
                AppInsightsShared.Tests.Configs.GetAiSecureConfig(false, false, false, false, false, false, true);

            string secureConfig = SerializationHelper.SerializeObject<AiSecureConfig>(aiSecureConfig);

            XrmFakedContext context = new XrmFakedContext();

            XrmFakedPluginExecutionContext xrmFakedPluginExecution = new XrmFakedPluginExecutionContext();
            Guid userId = Guid.Parse("9e7ec57b-3a08-4a41-a4d4-354d66f19b65");
            xrmFakedPluginExecution.InitiatingUserId = userId;
            xrmFakedPluginExecution.UserId = userId;
            xrmFakedPluginExecution.CorrelationId = Guid.Parse("15cc775b-9ebc-48d1-93a6-b0ce9c920b66");
            xrmFakedPluginExecution.MessageName = "update";
            xrmFakedPluginExecution.Mode = 1;
            xrmFakedPluginExecution.Depth = 1;
            xrmFakedPluginExecution.OrganizationName = "test.crm.dynamics.com";
            xrmFakedPluginExecution.Stage = 40;

            xrmFakedPluginExecution.InputParameters = GetInputParameters();

            xrmFakedPluginExecution.OutputParameters = new ParameterCollection();

            context.ExecutePluginWithConfigurations<LogEvent>(xrmFakedPluginExecution, null, secureConfig);

            Assert.IsTrue((bool)xrmFakedPluginExecution.OutputParameters["logsuccess"]);
        }

        [TestMethod]
        public void ActionEvent_Null_Measurement_Value_Test()
        {
            AiSecureConfig aiSecureConfig =
                AppInsightsShared.Tests.Configs.GetAiSecureConfig(false, false, false, false, false, false, true);

            string secureConfig = SerializationHelper.SerializeObject<AiSecureConfig>(aiSecureConfig);

            XrmFakedContext context = new XrmFakedContext();

            XrmFakedPluginExecutionContext xrmFakedPluginExecution = new XrmFakedPluginExecutionContext();

            xrmFakedPluginExecution.InputParameters = GetInputParameters();
            xrmFakedPluginExecution.InputParameters["measurementvalue"] = null;

            xrmFakedPluginExecution.OutputParameters = new ParameterCollection();

            context.ExecutePluginWithConfigurations<LogEvent>(xrmFakedPluginExecution, null, secureConfig);
        }

        [TestMethod]
        public void ActionEvent_Missing_Name_Test()
        {
            AiSecureConfig aiSecureConfig =
                AppInsightsShared.Tests.Configs.GetAiSecureConfig(false, false, false, false, false, false, true);

            string secureConfig = SerializationHelper.SerializeObject<AiSecureConfig>(aiSecureConfig);

            XrmFakedContext context = new XrmFakedContext();

            XrmFakedPluginExecutionContext xrmFakedPluginExecution = new XrmFakedPluginExecutionContext();

            xrmFakedPluginExecution.InputParameters = GetInputParameters();
            xrmFakedPluginExecution.InputParameters.Remove("name");

            xrmFakedPluginExecution.OutputParameters = new ParameterCollection();

            context.ExecutePluginWithConfigurations<LogEvent>(xrmFakedPluginExecution, null, secureConfig);

            Assert.IsFalse((bool)xrmFakedPluginExecution.OutputParameters["logsuccess"]);
            Assert.IsTrue(xrmFakedPluginExecution.OutputParameters["errormessage"].ToString() == "Name must be populated");
        }

        [TestMethod]
        public void ActionEvent_Null_Name_Test()
        {
            AiSecureConfig aiSecureConfig =
                AppInsightsShared.Tests.Configs.GetAiSecureConfig(false, false, false, false, false, false, true);

            string secureConfig = SerializationHelper.SerializeObject<AiSecureConfig>(aiSecureConfig);

            XrmFakedContext context = new XrmFakedContext();

            XrmFakedPluginExecutionContext xrmFakedPluginExecution = new XrmFakedPluginExecutionContext();

            xrmFakedPluginExecution.InputParameters = GetInputParameters();
            xrmFakedPluginExecution.InputParameters["name"] = null;
            ;

            xrmFakedPluginExecution.OutputParameters = new ParameterCollection();

            context.ExecutePluginWithConfigurations<LogEvent>(xrmFakedPluginExecution, null, secureConfig);

            Assert.IsFalse((bool)xrmFakedPluginExecution.OutputParameters["logsuccess"]);
            Assert.IsTrue(xrmFakedPluginExecution.OutputParameters["errormessage"].ToString() == "Name must be populated");
        }

        [TestMethod]
        public void ActionEvent_Invalid_Measurement_Name_Test()
        {
            AiSecureConfig aiSecureConfig =
                AppInsightsShared.Tests.Configs.GetAiSecureConfig(false, false, false, false, false, false, true);

            string secureConfig = SerializationHelper.SerializeObject<AiSecureConfig>(aiSecureConfig);

            XrmFakedContext context = new XrmFakedContext();

            XrmFakedPluginExecutionContext xrmFakedPluginExecution = new XrmFakedPluginExecutionContext();

            xrmFakedPluginExecution.InputParameters = GetInputParameters();
            xrmFakedPluginExecution.InputParameters["measurementname"] = "sisznxevfkzibsdtsvfwijucumzedrzauyzzyzqmrrmdwwdqugtiprgvgkmpokcoldnxcmlwywcuernvoobnfogzgjkbnsteycrvafpharlnylyvyigsnskjuwwqjeiudwkibztwzwwotfbaijxcqwwk";

            xrmFakedPluginExecution.OutputParameters = new ParameterCollection();

            context.ExecutePluginWithConfigurations<LogEvent>(xrmFakedPluginExecution, null, secureConfig);

            Assert.IsFalse((bool)xrmFakedPluginExecution.OutputParameters["logsuccess"]);
            Assert.IsTrue(xrmFakedPluginExecution.OutputParameters["errormessage"].ToString() == "Measurement name cannot exceed 150 characters");
        }

        private static ParameterCollection GetInputParameters()
        {
            return new ParameterCollection {
                new KeyValuePair<string, object>("name", "Hello from EventTest - 1"),
                new KeyValuePair<string, object>("measurementname", "TestMeasurement"),
                new KeyValuePair<string, object>("measurementvalue", 56f)
            };
        }
    }
}