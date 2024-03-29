using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using System.Text;

namespace ColorfullFlagIconDynamics365Engine
{
    /// <summary>
    /// Base class for all Plugins.
    /// </summary>  
    public class Plugin : IPlugin
    {
        protected const string PreImageName = "PreImage";
        protected const string PostImageName = "PostImage";
        
        protected class LocalPluginContext
        {
            internal IServiceProvider ServiceProvider
            {
                get;

                private set;
            }

            internal IOrganizationService OrganizationService
            {
                get;

                private set;
            }

            internal IOrganizationService AdminOrganizationService
            {
                get;

                private set;
            }

            internal IPluginExecutionContext PluginExecutionContext
            {
                get;

                private set;
            }

            internal ITracingService TracingService
            {
                get;

                private set;
            }

            internal StringBuilder TraceBuilder
            {
                get;

                private set;
            }

            private LocalPluginContext()
            {
            }

            internal LocalPluginContext(IServiceProvider serviceProvider)
            {
                if (serviceProvider == null)
                {
                    throw new ArgumentNullException("serviceProvider");
                }

                // Obtain the execution context service from the service provider.
                this.PluginExecutionContext = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

                // Obtain the tracing service from the service provider.
                this.TracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

                // Obtain the Organization Service factory service from the service provider
                IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));

                // Use the factory to generate the Organization Service.
                this.OrganizationService = factory.CreateOrganizationService(this.PluginExecutionContext.UserId);

                // Use the factory to generate the Admin Organization Service.
                this.AdminOrganizationService = factory.CreateOrganizationService(null);

                this.TraceBuilder = new StringBuilder();
            }

            internal void Trace(string message)
            {
                if (string.IsNullOrWhiteSpace(message) || this.TracingService == null)
                {
                    return;
                }

                if (this.PluginExecutionContext == null)
                {
                    this.TracingService.Trace(message);
                }
                else
                {
                    this.TracingService.Trace(
                        "{0}, Correlation Id: {1}, Initiating User: {2}",
                        message,
                        this.PluginExecutionContext.CorrelationId,
                        this.PluginExecutionContext.InitiatingUserId);
                }
            }

            internal void TracingError(Exception e, StringBuilder messageToLog)
            {
                this.TraceBuilder.Append(" | ");
                this.TraceBuilder.Append(this.PluginExecutionContext.PrimaryEntityName);
                this.TraceBuilder.Append(" - ");
                this.TraceBuilder.Append(this.PluginExecutionContext.Stage.ToString());
                this.TraceBuilder.Append(" - ");
                this.TraceBuilder.Append(this.PluginExecutionContext.MessageName);
                this.TraceBuilder.Append(" | ");
                this.TraceBuilder.Append(messageToLog);
                this.TraceBuilder.Append(" | ");
                this.TraceBuilder.Append(e.Message);
                this.TraceBuilder.Append(" | ");
                this.TraceBuilder.Append(e.StackTrace);
            }

            internal void TracingError(Exception e)
            {
                this.TraceBuilder.Append(" | ");
                this.TraceBuilder.Append(this.PluginExecutionContext.PrimaryEntityName);
                this.TraceBuilder.Append(" - ");
                this.TraceBuilder.Append(this.PluginExecutionContext.Stage.ToString());
                this.TraceBuilder.Append(" - ");
                this.TraceBuilder.Append(this.PluginExecutionContext.MessageName);
                this.TraceBuilder.Append(" | ");
                this.TraceBuilder.Append(e.Message);
                this.TraceBuilder.Append(" | ");
                this.TraceBuilder.Append(e.StackTrace);
            }

            #region image Helper
            internal Entity getPreEntityImage(string imageName)
            {
                VerifyPreImage(imageName);
                return this.PluginExecutionContext.PreEntityImages[imageName]; 
            }

            internal Entity getPostEntityImage(string imageName)
            {
                VerifyPostImage(imageName);
                return this.PluginExecutionContext.PostEntityImages[imageName]; 
            }

            protected void VerifyPreImage(string imageName)
            {
                VerifyImage(this.PluginExecutionContext.PreEntityImages, imageName, true);
            }

            protected void VerifyPostImage(string imageName)
            {
                VerifyImage(this.PluginExecutionContext.PostEntityImages, imageName, false);
            }

            private void VerifyImage(EntityImageCollection collection, string imageName, bool isPreImage)
            {
                if (!collection.Contains(imageName)
                    || collection[imageName] == null)
                {
                    throw new ArgumentNullException(imageName, string.Format("{0} {1} does not exist in this context", isPreImage ? "PreImage" : "PostImage", imageName));
                }
            }
            #endregion
        }

        private Collection<Tuple<int, string, string, Action<LocalPluginContext>>> registeredEvents;

        /// <summary>
        /// Gets the List of events that the plug-in should fire for. Each List
        /// Item is a <see cref="System.Tuple"/> containing the Pipeline Stage, Message and (optionally) the Primary Entity. 
        /// In addition, the fourth parameter provide the delegate to invoke on a matching registration.
        /// </summary>
        protected Collection<Tuple<int, string, string, Action<LocalPluginContext>>> RegisteredEvents
        {
            get
            {
                if (this.registeredEvents == null)
                {
                    this.registeredEvents = new Collection<Tuple<int, string, string, Action<LocalPluginContext>>>();
                }

                return this.registeredEvents;
            }
        }

        /// <summary>
        /// Gets or sets the name of the child class.
        /// </summary>
        /// <value>The name of the child class.</value>
        protected string ChildClassName
        {
            get;

            private set;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Plugin"/> class.
        /// </summary>
        /// <param name="childClassName">The <see cref=" cred="Type"/> of the derived class.</param>
        internal Plugin(Type childClassName)
        {
            this.ChildClassName = childClassName.ToString();
        }

        /// <summary>
        /// Executes the plug-in.
        /// </summary>
        /// <param name="serviceProvider">The service provider.</param>
        /// <remarks>
        /// For improved performance, Microsoft Dynamics CRM caches plug-in instances. 
        /// The plug-in's Execute method should be written to be stateless as the constructor 
        /// is not called for every invocation of the plug-in. Also, multiple system threads 
        /// could execute the plug-in at the same time. All per invocation state information 
        /// is stored in the context. This means that you should not use global variables in plug-ins.
        /// </remarks>
        public void Execute(IServiceProvider serviceProvider)
        {
            if (serviceProvider == null)
            {
                throw new ArgumentNullException("serviceProvider");
            }

            // Construct the Local plug-in context.
            LocalPluginContext localcontext = new LocalPluginContext(serviceProvider);

            localcontext.Trace(string.Format(CultureInfo.InvariantCulture, "Entered {0}.Execute()", this.ChildClassName));

            try
            {
                // Iterate over all of the expected registered events to ensure that the plugin
                // has been invoked by an expected event
                // For any given plug-in event at an instance in time, we would expect at most 1 result to match.
                Action<LocalPluginContext> entityAction =
                    (from a in this.RegisteredEvents
                     where (
                     a.Item1 == localcontext.PluginExecutionContext.Stage &&
                     a.Item2 == localcontext.PluginExecutionContext.MessageName &&
                     (string.IsNullOrWhiteSpace(a.Item3) ? true : a.Item3 == localcontext.PluginExecutionContext.PrimaryEntityName)
                     )
                     select a.Item4).FirstOrDefault();

                if (entityAction != null)
                {
                    localcontext.Trace(string.Format(
                        CultureInfo.InvariantCulture,
                        "{0} is firing for Entity: {1}, Message: {2}",
                        this.ChildClassName,
                        localcontext.PluginExecutionContext.PrimaryEntityName,
                        localcontext.PluginExecutionContext.MessageName));

                    entityAction.Invoke(localcontext);

                    // now exit - if the derived plug-in has incorrectly registered overlapping event registrations,
                    // guard against multiple executions.
                    return;
                }
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                localcontext.Trace(string.Format(CultureInfo.InvariantCulture, "Exception: {0}", e.ToString()));

                // Handle the exception.
                throw;
            }
            finally
            {
                localcontext.Trace(string.Format(CultureInfo.InvariantCulture, "Exiting {0}.Execute()", this.ChildClassName));
            }
        }
    }
}