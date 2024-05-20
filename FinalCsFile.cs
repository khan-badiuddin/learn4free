using System;
using System.Diagnostics;
using System.ServiceModel;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace CreateNotes
{
    public class CreateNotes : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            // Obtain the tracing service
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                tracingService.Trace("Plugin execution started.");

                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                #region On Action Create and Update

                if (context.PrimaryEntityName.ToLower() == "xom_guyanalearningapplicationaction" && (context.MessageName.ToLower() == "create" || context.MessageName.ToLower() == "update"))
                {
                    tracingService.Trace("Plugin triggered for Action entity.");

                    if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                    {
                        Entity targetEntity = (Entity)context.InputParameters["Target"];

                        if (targetEntity.Attributes.Contains("llm_notes"))
                        {
                            tracingService.Trace("Notes description field is being updated.");

                            string notesDescription = StripHTML(targetEntity.GetAttributeValue<string>("llm_notes"));

                            #region On Action Create

                            if (context.MessageName.ToLower() == "create")
                            {
                                tracingService.Trace("Creating new annotation for Action.");

                                Entity noteEntity = new Entity("annotation");
                                noteEntity["objectid"] = new EntityReference("xom_guyanalearningapplicationaction", targetEntity.Id);
                                noteEntity["notetext"] = notesDescription;
                                service.Create(noteEntity);

                                tracingService.Trace("New annotation created for Action.");
                            }

                            #endregion


                            #region On Action Update

                            else if (context.MessageName.ToLower() == "update")
                            {
                                tracingService.Trace("Updating existing annotation for Action.");

                                Entity latestNote = GetLatestNoteForAccount(targetEntity.Id, service);
								
                                if (targetEntity.GetAttributeValue<string>("llm_notes") != "N/A")
                                {
                                    if (latestNote != null)
                                    {
                                        tracingService.Trace("Latest note found for Action.");

                                        string previousNoteText = latestNote.GetAttributeValue<string>("notetext");

                                        if (notesDescription != previousNoteText)
                                        {
                                            tracingService.Trace("Notes description is changed. Updating annotation.");

                                            latestNote["notetext"] = notesDescription;
                                            service.Update(latestNote);

                                            tracingService.Trace("Annotation updated for Action.");
                                        }
                                        else
                                        {
                                            tracingService.Trace("Notes description is not changed. No update needed.");
                                        }


                                    }
                                    else
                                    {
                                        tracingService.Trace("No existing note found for Action. Creating new annotation.");

                                        Entity noteEntity = new Entity("annotation");
                                        noteEntity["objectid"] = new EntityReference("xom_guyanalearningapplicationaction", targetEntity.Id);
                                        noteEntity["notetext"] = notesDescription;
                                        service.Create(noteEntity);

                                        tracingService.Trace("New annotation created for Action.");
                                    }
                                }
                                else
                                {
                                    tracingService.Trace("Accounts descrition is N/A");
                                }
                            }

                            #endregion
                        }
                    }
                }

                #endregion

                #region On Annotation Create and Update and Delete

                if (context.PrimaryEntityName.ToLower() == "annotation" && (context.MessageName.ToLower() == "create" || context.MessageName.ToLower() == "update"))
                {
                    tracingService.Trace("Plugin triggered for Annotation entity.");

                    if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                    {
                        Entity targetEntity = (Entity)context.InputParameters["Target"];

                        if (targetEntity.Attributes.Contains("notetext"))
                        {
                            tracingService.Trace("Note text field is being updated.");

                            Guid actionId;

                            #region On Annotation Create

                            if (context.MessageName.ToLower() == "create")
                            {
                                tracingService.Trace("Creating a new annotation.");

                                EntityReference objectRef = targetEntity.GetAttributeValue<EntityReference>("objectid");
								
                                if (objectRef == null || objectRef.LogicalName.ToLower() != "xom_guyanalearningapplicationaction")
                                {
                                    tracingService.Trace("Annotation is not related to an Action.");
                                    return;
                                }
                                actionId = objectRef.Id;
                            }

                            #endregion

                            #region On Annotation Update

                            else
                            {
                                tracingService.Trace("Updating an existing annotation.");

                                Entity preImage = context.PreEntityImages.Contains("PreImage") ? context.PreEntityImages["PreImage"] : null;
                                if (preImage == null)
                                {
                                    tracingService.Trace("PreImage is not available.");
                                    return;
                                }
                                actionId = ((EntityReference)preImage["objectid"]).Id;
                            }

                            

                            Entity action = service.Retrieve("xom_guyanalearningapplicationaction", actionId, new ColumnSet("llm_notes"));

                            string currentDescription = action.GetAttributeValue<string>("llm_notes");
                            string newNoteText = targetEntity.GetAttributeValue<string>("notetext");
                            string strippedNoteText = StripHTML(newNoteText);

                            if (currentDescription != strippedNoteText)
                            {
                                tracingService.Trace("Action description is different from note text. Updating action description.");

                                action["llm_notes"] = strippedNoteText;
                                service.Update(action);

                                tracingService.Trace("Action description updated.");
                            }
                            else
                            {
                                tracingService.Trace("Action description is the same as note text. No update needed.");
                            }

                            #endregion
                        }
                    }
                }

                #region On Annotation Delete

                if (context.MessageName.ToLower() == "delete" && context.PrimaryEntityName.ToLower() == "annotation")
                {
                    tracingService.Trace("Plugin triggered for Annotation deletion.");

                    Entity preImage = context.PreEntityImages.Contains("PreImage") ? context.PreEntityImages["PreImage"] as Entity : null;
					
                    if (preImage == null || !preImage.Contains("objectid") || preImage.GetAttributeValue<EntityReference>("objectid").LogicalName.ToLower() != "action")
                    {
                        tracingService.Trace("PreImage is not available or does not contain valid data.");
                        return;
                    }

                    Guid actionId = preImage.GetAttributeValue<EntityReference>("objectid").Id;

                    EntityCollection previousNotes = GetPreviousNotesForAccount(actionId, preImage.Id, service);

                    if (previousNotes.Entities.Count > 0)
                    {
                        Entity previousNote = previousNotes.Entities[0];

                        tracingService.Trace("Previous note found. Updating Action description.");

                        Entity actionEntity = service.Retrieve("xom_guyanalearningapplicationaction", actionId, new ColumnSet("llm_notes"));

                        string previousNoteText = previousNote.GetAttributeValue<string>("notetext");
                        string strippedPreviousNoteText = StripHTML(previousNoteText);

                        actionEntity["llm_notes"] = strippedPreviousNoteText;
                        service.Update(actionEntity);

                        tracingService.Trace("Action description updated with previous note text.");
                    }
                    else
                    {
                        Entity ActiontoUpdate = new Entity("xom_guyanalearningapplicationaction");
                        ActiontoUpdate.Id = actionId;
                        ActiontoUpdate["llm_notes"] = "N/A";

                        tracingService.Trace("Plugin will now Update Accoutn with N/A");
                        service.Update(ActiontoUpdate);
                        tracingService.Trace("Plugin entered Else condition of note delte when no note is present");                      
                    }

                    #endregion

                }

                #endregion

            }
            catch (Exception ex)
            {
                tracingService.Trace("Error occurred: " + ex.ToString());
                throw new InvalidPluginExecutionException("An error occurred in the CreateNotes plugin.", ex);
            }
            finally
            {
                tracingService.Trace("Plugin execution completed.");
            }
        }

        public static string StripHTML(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return string.Empty;

            // Remove HTML tags and &nbsp;
            string output = Regex.Replace(input, "<.*?>", String.Empty);
            output = Regex.Replace(output, "&nbsp;", " ", RegexOptions.IgnoreCase);
            return output;
        }

        private Entity GetLatestNoteForAccount(Guid actionId, IOrganizationService service)
        {
            QueryExpression query = new QueryExpression("annotation");
            query.ColumnSet = new ColumnSet("annotationid", "notetext", "objectid");
            query.Criteria.AddCondition("objectid", ConditionOperator.Equal, actionId);
            query.AddOrder("modifiedon", OrderType.Descending);
            query.TopCount = 1;

            EntityCollection notes = service.RetrieveMultiple(query);
            if (notes.Entities.Count > 0)
            {
                return notes.Entities[0];
            }
            return null;
        }

        private EntityCollection GetPreviousNotesForAccount(Guid actionId, Guid latestNoteId, IOrganizationService service)
        {
            QueryExpression query = new QueryExpression("annotation");
            query.ColumnSet = new ColumnSet("annotationid", "notetext", "objectid");
            query.Criteria.AddCondition("objectid", ConditionOperator.Equal, actionId);
            query.Criteria.AddCondition("annotationid", ConditionOperator.NotEqual, latestNoteId);
            query.AddOrder("modifiedon", OrderType.Descending);
            query.TopCount = 1;

            return service.RetrieveMultiple(query);
        }
    }
}
