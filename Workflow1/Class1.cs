// <copyright file="Class1.cs" company="">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <author></author>
// <date>4/22/2015 5:42:50 PM</date>
// <summary>Implements the Class1 Workflow Activity.</summary>
namespace CustomWorkflowActivities.Workflow1
{
    using System;
    using System.Activities;
    using System.ServiceModel;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Workflow;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Query;


    public sealed class Class1 : CodeActivity
    {
        /// <summary>
        /// Executes the workflow activity.
        /// </summary>
        /// <param name="executionContext">The execution context.</param>
        protected override void Execute(CodeActivityContext executionContext)
        {

            // Workflow plugin
            // Given an Opportunity, Quote, Order or Invoice
            // Output a string for use in email or other sources.

            // April 2015
            // Bob Hatcher bob.hatcher@gmail.com
            // http://ms-crm.guru

            // Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            if (tracingService == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve tracing service.");
            }

            tracingService.Trace("Entered Class1.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}",
                executionContext.ActivityInstanceId,
                executionContext.WorkflowInstanceId);

            // Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

            if (context == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");
            }

            tracingService.Trace("Class1.Execute(), Correlation Id: {0}, Initiating User: {1}",
                context.CorrelationId,
                context.InitiatingUserId);

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                EntityReference entity = null;
                String baseEntity = ((String)this.TypeInput.Get(executionContext)).ToLower();
                String schemaName = "";
                // in case the user enters order and not salesorder
                if (baseEntity.Equals("order"))
                    baseEntity = "salesorder";
                    
                // based on the input, create the source entity and set the child entity type for use later.
                if (baseEntity.Equals("opportunity")){
                    entity = this.OpportunityInput.Get(executionContext);
                    schemaName = "opportunityproduct";
                }
                if (baseEntity.Equals("salesorder")) {
                    entity = this.OrderInput.Get(executionContext);
                    schemaName = "salesorderdetail";
                }
                if (baseEntity.Equals("quote"))
                {
                    entity = this.QuoteInput.Get(executionContext);
                    schemaName = "quotedetail";
                }
                if (baseEntity.Equals("invoice"))
                {
                    entity = this.InvoiceInput.Get(executionContext);
                    schemaName = "invoicedetail";
                }

                if (entity == null)
                    throw new InvalidOperationException("Entity has not been specified", new ArgumentNullException("Value " + baseEntity + " is invalid"));

                // Other input parameters
                String currencySymbol = "$";
                if ((this.CurrencySymbolInput.Get(executionContext) != null) && (!String.IsNullOrEmpty((String)this.CurrencySymbolInput.Get(executionContext))))
                    currencySymbol = (String)this.CurrencySymbolInput.Get(executionContext);

                int quantityPrecision =  (int)this.QuantityPrecisionInput.Get(executionContext);

                int currencyPrecision = (int)this.CurrencyPrecisionInput.Get(executionContext);

                bool bundleHeaders = (bool)this.BundleHeadersInput.Get(executionContext);

                //Retrieve the CrmService so that we can retrieve the source entity
                IWorkflowContext wfContext = executionContext.GetExtension<IWorkflowContext>();
                IOrganizationServiceFactory wfServiceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
                IOrganizationService wfService = wfServiceFactory.CreateOrganizationService(context.InitiatingUserId);

                //Retrieve the Opportunity Entity since we need its guid
                Entity sourceEntity;
                {
                    //Create a request
                    RetrieveRequest retrieveRequest = new RetrieveRequest();
                    retrieveRequest.ColumnSet = new ColumnSet(new string[] { baseEntity+"id" });
                    retrieveRequest.Target = entity;

                    //Execute the request
                    RetrieveResponse retrieveResponse = (RetrieveResponse)wfService.Execute(retrieveRequest);

                    //Retrieve the Loan Application Entity
                    sourceEntity = retrieveResponse.Entity as Entity;
                }

                if (!sourceEntity.Contains(baseEntity + "id"))
                    return;

                String sourceEntityId = (sourceEntity.GetAttributeValue<Guid>(baseEntity + "id")).ToString();


                // This fetch gets all related entities.
                // schemaname, as specified above, is the child type (opportunityproduct, quotedetail, invoicedetail, orderdetail)
                String fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='" + schemaName + @"'>
                <attribute name='productid' />
                <attribute name='productdescription' />
                <attribute name='priceperunit' />
                <attribute name='extendedamount' />
                <attribute name='quantity' />
                <attribute name='extendedamount' />
                <attribute name='" + schemaName + @"id' />
                <order attribute='productid' descending='false' />";

                // If the user asked that the headers be included, include only Headers and not the detailed line items.
                if (bundleHeaders)
                    fetch += @"<filter type='and'>
                      <filter type='or'>
                        <condition attribute='producttypecode' operator='in'>
                          <value>1</value>
                          <value>2</value>
                        </condition>
                        <condition attribute='producttypecode' operator='null' />
                      </filter>
                    </filter>";
                else
                    fetch += @"<filter type='and'>
                      <filter type='or'>
                        <condition attribute='producttypecode' operator='in'>
                          <value>1</value>
                          <value>4</value>
                          <value>3</value>
                        </condition>
                        <condition attribute='producttypecode' operator='null' />
                      </filter>
                    </filter>";
                fetch += @"
                <link-entity name='" + baseEntity + @"' from='" + baseEntity + @"id' to='" + baseEntity + @"id' alias='ab'>
                    <filter type='and'>
                    <condition attribute='" + baseEntity + @"id' operator='eq' uitype='" + baseEntity + @"' value='{" + sourceEntityId + @"}' />
                    </filter>
                </link-entity>
                </entity>
                </fetch>";
                EntityCollection result = wfService.RetrieveMultiple(new FetchExpression(fetch));

                String outputString = "This " + baseEntity + " contains " + result.Entities.Count + " line items.\n" ;

                // Now iterate over the line items and build an output string.

                if (result != null && result.Entities.Count > 0)
                {
                    // In this section _entity is the returned one
                    foreach (Entity _entity in result.Entities)
                    {
                        //outputString = "";
                        if (_entity.Contains("quantity"))
                            outputString += " Qty " + Convert.ToString(Math.Round(Convert.ToDecimal(_entity.Attributes["quantity"]),quantityPrecision)) + " - ";
                        else
                            outputString += " Qty unknown - ";
                        //(String)((AliasedValue)_entity["cr.jr_streetaddress"]).Value)

                        // Assume it's going to have a productid or a productdescription.
                        if (_entity.Contains("productid"))
                            outputString += " " + Convert.ToString(((EntityReference)_entity.Attributes["productid"]).Name) + " - ";
                        else
                            outputString += " Write-In Product: ";

                        if (_entity.Contains("productdescription"))
                            outputString += " " + _entity.Attributes["productdescription"] + " - ";
                        // no else

                        if (_entity.Contains("priceperunit"))
                            outputString += currencySymbol + Convert.ToString(Math.Round(Convert.ToDecimal(_entity.GetAttributeValue<Money>("priceperunit").Value),currencyPrecision)) + " per unit ";
                        else
                            outputString += currencySymbol + "0 per unit ";

                        if (_entity.Contains("extendedamount"))
                            outputString += " -- " + currencySymbol + Convert.ToString(Math.Round(Convert.ToDecimal(_entity.GetAttributeValue<Money>("extendedamount").Value), currencyPrecision)) + " total after discounts ";
                        else
                            outputString += currencySymbol + " (no total) ";

                        outputString += "\n";
                            
                    }

                }

                // Output results.
                this.ConsolidatedProductString.Set(executionContext, outputString);
                
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                tracingService.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw;
            }

            tracingService.Trace("Exiting Class1.Execute(), Correlation Id: {0}", context.CorrelationId);
        }
        //Define the properties
        [Input("Opportunity")]
        [ReferenceTarget("opportunity")]
        public InArgument<EntityReference> OpportunityInput { get; set; }

        [Input("Order")]
        [ReferenceTarget("salesorder")]
        public InArgument<EntityReference> OrderInput { get; set; }

        [Input("Quote")]
        [ReferenceTarget("quote")]
        public InArgument<EntityReference> QuoteInput { get; set; }

        [Input("Invoice")]
        [ReferenceTarget("invoice")]
        public InArgument<EntityReference> InvoiceInput { get; set; }

        [Input("Type")]
        public InArgument<String> TypeInput { get; set; }

        [Input("Currency Precision")]
        public InArgument<int> CurrencyPrecisionInput { get; set; }

        [Input("Quantity Precision")]
        public InArgument<int> QuantityPrecisionInput { get; set; }

        [Input("Currency Symbol")]
        public InArgument<String> CurrencySymbolInput { get; set; }

        [Input("Bundle Headers")]
        public InArgument<bool> BundleHeadersInput { get; set; }

        [Output("Consolidated Line Item Output")]
        [AttributeTarget("email", "description")]
        public OutArgument<String> ConsolidatedProductString { get; set; }
    }
}