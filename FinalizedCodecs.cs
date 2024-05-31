using System;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
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
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                tracingService.Trace("Plugin execution started.");

                #region Action entity Operations

                #region On Action Create

                if (context.PrimaryEntityName.ToLower() == "xom_guyanalearningapplicationaction" && (context.MessageName.ToLower() == "create"))
                {
                    tracingService.Trace("Plugin Entered Create Condition.");

                    if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                    {
                        Entity targetEntity = (Entity)context.InputParameters["Target"];

                       
                        if (targetEntity.Attributes.Contains("llm_notes")) 
                        {
                            tracingService.Trace("Actions Notes description Contains Data.");

                            string notesDescription = StripHTML(targetEntity.GetAttributeValue<string>("llm_notes"));

                            tracingService.Trace("Creating new annotation for Action.");

                            Entity noteEntity = new Entity("annotation");
                            noteEntity["objectid"] = new EntityReference("xom_guyanalearningapplicationaction", targetEntity.Id);
                            noteEntity["notetext"] = notesDescription;
                            service.Create(noteEntity);

                            tracingService.Trace("New annotation created for Action with ActionDescription Value.");
                        }
                        else // Action Description Doesn't have Value
                        {
                            tracingService.Trace("Actions Notes Description Does not Contain Data, So no Annotation is Created.");
                        }
                    }
                }
                #endregion


                #region On Action Update

                if (context.PrimaryEntityName.ToLower() == "xom_guyanalearningapplicationaction" && context.MessageName.ToLower() == "update")
                {
                    if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                    {
                        Entity targetEntity = (Entity)context.InputParameters["Target"];

                        tracingService.Trace("Plugin has entered Update Condition.");

                        Entity latestNote = GetLatestNoteForAction(targetEntity.Id, service);
                        string notesDescription = StripHTML(targetEntity.GetAttributeValue<string>("llm_notes"));


                        if (latestNote == null)// This Condition Checks If no Existing Annotation is Associated with Action.
                        {
                            // now will check if the Action Description is not equal to N/A so that we can create a Annotation.
                            if (targetEntity.GetAttributeValue<string>("llm_notes") != "N/A") 
                            {
                                tracingService.Trace("Action Description is not N/A.");

                                Entity annotation = new Entity("annotation");
                                annotation["notetext"] = notesDescription;
                                annotation["objectid"] = new EntityReference("xom_guyanalearningapplicationaction", targetEntity.Id);

                                service.Create(annotation);
                                tracingService.Trace("New Notes Created on Update of Action NoteDescription");
                            }
                            else// because the Action Description is N/A it should not create a Annotation
                            {
                                tracingService.Trace("Actions description is N/A,so no new annotation is Created.");
                            }
                        }
                        else
                        {
                            string AnnotationNoteText = StripHTML(latestNote.GetAttributeValue<string>("notetext"));
                                                       
                                if (notesDescription != AnnotationNoteText)
                                {
                                    tracingService.Trace(" Action Notes description is changed and now contains Value. Creating annotation.");

                                    Entity notes = new Entity("annotation");
                                    notes["subject"] = "On Action Update";
                                    notes["notetext"] = notesDescription;
                                    notes["objectid"] = new EntityReference("xom_guyanalearningapplicationaction", targetEntity.Id);
                                    service.Create(notes);
                                    tracingService.Trace("Annotation Created on Update of Action NoteDescription");
                                }
                                else
                                {
                                    tracingService.Trace("Already an annotation with Similar notetext Exist");
                                }
                        }
                    }
                }
                #endregion

                #endregion


                #region Annotation Entity Operations


                if (context.PrimaryEntityName.ToLower() == "annotation" && (context.MessageName.ToLower() == "create" || context.MessageName.ToLower() == "update"))
                {
                    tracingService.Trace("Plugin triggered for Annotation entity.");

                    if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                    {
                        Entity targetEntity = (Entity)context.InputParameters["Target"];

                        if (targetEntity.Attributes.Contains("notetext"))
                        {
                            tracingService.Trace("NoteText field has Data.");

                            Guid actionId;

                            #region On Annotation Create

                            if (context.MessageName.ToLower() == "create")
                            {
                                tracingService.Trace("Plugin entered annotation Create Step.");

                                EntityReference objectRef = targetEntity.GetAttributeValue<EntityReference>("objectid");

                                if (objectRef == null || objectRef.LogicalName.ToLower() != "xom_guyanalearningapplicationaction")
                                {
                                    tracingService.Trace("Annotation is not related to an action.");
                                    return;
                                }
                                actionId = objectRef.Id;
                            }
                            #endregion

                            #region On Annotation Update

                            else
                            {
                                tracingService.Trace("Plugin Entered Annotation Update Step.");

                                Entity preImage = context.PreEntityImages.Contains("PreImage") ? context.PreEntityImages["PreImage"] : null;

                                if (preImage == null)
                                {
                                    tracingService.Trace("PreImage is not available.");

                                    return;
                                }
                                actionId = ((EntityReference)preImage["objectid"]).Id;
                            }

                            #endregion

                            Entity action = service.Retrieve("xom_guyanalearningapplicationaction", actionId, new ColumnSet("llm_notes"));

                            string currentDescription = action.GetAttributeValue<string>("llm_notes");
                            string CurrentDesc = StripHTML(currentDescription);

                            string newNoteText = targetEntity.GetAttributeValue<string>("notetext");
                            string strippedNoteText = StripHTML(newNoteText);

                            if (CurrentDesc != strippedNoteText)
                            {
                                tracingService.Trace("Action description is different from notetext. need to Update action description.");

                                action["llm_notes"] = strippedNoteText;

                                service.Update(action);

                                tracingService.Trace("Action description updated.");
                            }
                            else
                            {
                                tracingService.Trace("Action description is the same as note text. No update needed.");
                            }
                        }
                    }
                }

                #endregion

                #region annoation Delete Operation


                if (context.MessageName.ToLower() == "delete" && context.PrimaryEntityName.ToLower() == "annotation")
                {
                    tracingService.Trace("Plugin triggered for Annotation deletion.");

                    Entity preImage = context.PreEntityImages.Contains("PreImage") ? context.PreEntityImages["PreImage"] as Entity : null;

                    if (preImage == null || !preImage.Contains("objectid") || preImage.GetAttributeValue<EntityReference>("objectid").LogicalName.ToLower() != "xom_guyanalearningapplicationaction")
                    {
                        tracingService.Trace("PreImage is not available or does not contain valid data.");
                        return;
                    }

                    Guid actionId = preImage.GetAttributeValue<EntityReference>("objectid").Id;

                    EntityCollection previousNotes = GetPreviousNotesForAction(actionId, preImage.Id, service);

                    if (previousNotes.Entities.Count > 0)
                    {
                        Entity previousNote = previousNotes.Entities[0];

                        tracingService.Trace("Previous note found. Updating action description.");

                        Entity actionEntity = service.Retrieve("xom_guyanalearningapplicationaction", actionId, new ColumnSet("llm_notes"));

                        string previousNoteText = previousNote.GetAttributeValue<string>("notetext");
                        string strippedPreviousNoteText = StripHTML(previousNoteText);

                        actionEntity["llm_notes"] = strippedPreviousNoteText;

                        service.Update(actionEntity);

                        tracingService.Trace("Action description updated with previous note text.");
                    }
                    else
                    {
                        Entity Actoupdate = new Entity("xom_guyanalearningapplicationaction");
                        Actoupdate.Id = actionId;
                        Actoupdate["llm_notes"] = "N/A";

                        tracingService.Trace("Updating Action with N/A, as no Annotation is Left.");

                        service.Update(Actoupdate);

                        tracingService.Trace("Action Description Updated with N/A.");
                    }

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

        private Entity GetLatestNoteForAction(Guid actionId, IOrganizationService service)
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
        private EntityCollection GetPreviousNotesForAction(Guid actionId, Guid latestNoteId, IOrganizationService service)
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
