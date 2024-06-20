#region C# Code
$OlderThan2017Cs = @'
	using System;
	using System.Collections.Generic;
	using System.Linq;

	// Vault SDK
	using vltobj = Autodesk.Connectivity.WebServices;
	using vltutil = Autodesk.Connectivity.WebServicesTools;

	namespace Autodesk.Vault.SDKTools
	{
		/// <summary>
		/// Helper class to do a property sync using filestore service.
		/// This code has been tested with the 2016 Vault API. It should also 
		/// work for 2015-R2 API, but it will not work with 2015 or earlier API.
		/// </summary>
		/// <author>Dave Mink</author>
		public class PropertySync
		{
			// property def cache
			// Needed in order to determine how to parse the invariant string value we get from GetComponentProperties
			private Dictionary<string, Dictionary<long, vltobj.PropDef>> m_propDefsByEntityClassAndId = new Dictionary<string, Dictionary<long, vltobj.PropDef>>();

			public PropertySync(vltutil.WebServiceManager svcmgr)
			{
				// properties can come from files or items.
				foreach(string entityClass in new string[] { "FILE", "ITEM" })
				{
					m_propDefsByEntityClassAndId.Add(
						entityClass,
						svcmgr.PropertyService.GetPropertyDefinitionsByEntityClassId(entityClass).ToDictionary(pd => pd.Id)
						);
				}
			}

			/// <summary>
			/// Sync properties of a file.
			/// </summary>
			/// <param name="svcmgr">a WebServiceManager</param>
			/// <param name="file">the file you would like to sync</param>
			/// <param name="comment">the comment for the new version (if a property sync was performed)</param>
			/// <param name="allowSync">if the local filestore doesn't have the file, get it from another filestore</param>
			/// <param name="writeResults">see FilestoreService.CopyFile method</param>
			/// <param name="cloakedEntityClasses">if you can't read an entity where properties would come from, its entity class is returned here</param>
			/// <param name="force">skip check for equivalence and always do a sync, creating a new version</param>
			/// <returns>the file returned from the checkin, same as the input if no property sync is done</returns>
			public vltobj.File SyncProperties(vltutil.WebServiceManager svcmgr, vltobj.File file, string comment, bool allowSync, out vltobj.PropWriteResults writeResults, out string[] cloakedEntityClasses, bool force=false)
			{
				// NOTE: we could have held onto svcmgr passed to constructor, 
				// but holding onto to something like that is generally a bad idea.

				// clear output parameters so we don't have to worry about that for each possible return condition.
				writeResults = null;
				cloakedEntityClasses = null;

				// first check for property compliance failures.
				// We don't need to sync unless there are equivalence failures.
				if (!force) // skip this fast-out if we are forcing a sync.
				{
					// NOTE: if we synced props to multiple files at a time, we could get compliance failures for all of them in one call.
					vltobj.PropCompFail[] complianceFailures = svcmgr.PropertyService.GetPropertyComplianceFailuresByEntityIds(
						"FILE", new long[] { file.Id }, /*filterPending*/true
						);
					if (complianceFailures == null 
						|| complianceFailures.Sum(cf => (cf.PropEquivFailArray != null ? cf.PropEquivFailArray.Length : 0)) == 0
						)
					{
						// nothing to do!
						return file;
					}
				}

				// checkout file without downloading it.
				// this will fail if file is already checked out.
				vltobj.ByteArray downloadTicket;
				file = svcmgr.DocumentService.CheckoutFile(
					file.Id, vltobj.CheckoutFileOptions.Master,
					/*machine*/Environment.MachineName, /*localPath*/string.Empty, comment,
					out downloadTicket
					);

				try // if anything goes wrong from here on out, undo the checkout
				{
					// get component properties.
					// NOTE: Behavior of this API changed significantly in 2015-R2.
					//       This will not work correctly in 2015 or earlier!
					// NOTE: a null component UID means to get write-back properties for root component in the file.
					// WARNING: we can't sync component level properties without CAD!
					vltobj.CompProp[] compProps = svcmgr.DocumentService.GetComponentProperties(file.Id, /*compUID*/null);

					// return cloaked entity classes thru out parameter.
					// NOTE: a propDefId of -1 indicates get couldn't get properties from an inaccessible entity.
					cloakedEntityClasses = compProps.Where(p => p.PropDefId < 0).Select(p => p.EntClassId).ToArray();
					if (cloakedEntityClasses != null && cloakedEntityClasses.Length > 0)
					{
						// don't proceed since we don't have the permissions to write back 
						// everything that is necessary to clear the failures.
						return svcmgr.DocumentService.UndoCheckoutFile(file.MasterId, out downloadTicket);
					}

					// filter so we only keep values from non-cloaked entities
					// NOTE: this is unnecessary as long as we bail out if there are cloaked entities involved.
					compProps = compProps.Where(p => p.PropDefId > 0).ToArray();
					
					// if there is nothing to write back, bail out.
					// We shouldn't have made it this far if this was the case;
					// but you never can be too sure!
					if (compProps == null || compProps.Length == 0)
					{
						// nothing to do, undo checkout and return
						return svcmgr.DocumentService.UndoCheckoutFile(file.MasterId, out downloadTicket);
					}

					// convert CompProp array to PropWriteReq array.
					//	It really sucks that the Value is a string - it should be an object.
					//	Since it is an invariant string, we need to parse it using invariant locales and convert dates from UTC.
					//	To do that we need to know what the type is, which we can get from the PropertyDef for the PropDefId.
					//	Since we don't want to get property defs everytime this is called, we use a cache.
					//	This should be fixed in a future release.
					vltobj.PropWriteReq[] writeProps = compProps.Select(
						p => new vltobj.PropWriteReq()
						{
							Moniker = p.Moniker,
							CanCreate = p.CreateNew,
							Val = ConvertInvariantStringToObject(p.Value, m_propDefsByEntityClassAndId[p.EntClassId][p.PropDefId].Typ)
						}
						).ToArray();

					// use CopyFile to copy existing resource and write the properties.
					byte[] uploadTicket = svcmgr.FilestoreService.CopyFile(
						downloadTicket.Bytes, allowSync, writeProps,
						out writeResults
						);

					// get child file associations so we can preserve them.
					// NOTE: if we synced props to multiple files at a time, we could get file associations for all of them in one call.
					vltobj.FileAssocLite[] childAssocs = svcmgr.DocumentService.GetFileAssociationLitesByIds(
						new long[] { file.Id },
						vltobj.FileAssocAlg.Actual, // preserve the associations provided by CAD add-in
						/*parentAssociationType*/vltobj.FileAssociationTypeEnum.None, /*parentRecurse*/false,
						/*childAssociationType*/vltobj.FileAssociationTypeEnum.All, /*childRecurse*/false,
						/*includeLibraryFiles*/true,
						/*includeRelatedDocuments*/false,
						/*includeHidden*/true
						);
					// convert FileAssocLite array to FileAssocParam array
					vltobj.FileAssocParam[] associations = childAssocs.Select(
						a => new vltobj.FileAssocParam()
							{
								Typ = a.Typ, CldFileId = a.CldFileId,
								Source = a.Source, RefId = a.RefId, 
								ExpectedVaultPath = a.ExpectedVaultPath
							}
						).ToArray();
					// checkin file
					file = svcmgr.DocumentService.CheckinUploadedFile(
						file.MasterId,
						comment, /*keepCheckedOut*/false, /*lastWrite*/DateTime.Now,
						associations,
						/*bom*/null, /*copyBom*/true, // preserve any BOM
						file.Name, file.FileClass, file.Hidden, // preserve these attributes
						new vltobj.ByteArray() { Bytes = uploadTicket }
						);
				}
				finally
				{
					// if we got here and file is still checked-out, 
					// something went wrong so undo the checkout.
					if (file.CheckedOut)
						file = svcmgr.DocumentService.UndoCheckoutFile(file.MasterId, out downloadTicket);
				}

				return file;
			}

			private static object ConvertInvariantStringToObject(string value, vltobj.DataType typ)
			{
				if (value == null) return null; // quick exit on null

				// dates need to parsed with invariant culture and converted to local time
				if (typ == vltobj.DataType.DateTime)
					return DateTime.Parse(value, System.Globalization.CultureInfo.InvariantCulture).ToLocalTime();
				// numerics need to be parsed with invariant culture
				else if (typ == vltobj.DataType.Numeric)
					return Double.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
				// bools are straight forward
				else if (typ == vltobj.DataType.Bool)
					return Boolean.Parse(value);
				else // don't do anything for string or image types (image types should not be seen here)
					return value;
			}
		}
	}

'@

$2017AndNewerCs = @'
	using System;
	using System.Collections.Generic;
	using System.Linq;

	// Vault SDK
	using vltobj = Autodesk.Connectivity.WebServices;
	using vltutil = Autodesk.Connectivity.WebServicesTools;

	namespace Autodesk.Vault.SDKTools
	{
		/// <summary>
		/// Helper class to do a property sync using filestore service.
		/// This code has been tested with the 2016 Vault API. It should also 
		/// work for 2015-R2 API, but it will not work with 2015 or earlier API.
		/// </summary>
		/// <author>Dave Mink</author>
		public class PropertySync
		{
			// property def cache
			// Needed in order to determine how to parse the invariant string value we get from GetComponentProperties
			private Dictionary<string, Dictionary<long, vltobj.PropDef>> m_propDefsByEntityClassAndId = new Dictionary<string, Dictionary<long, vltobj.PropDef>>();

			public PropertySync(vltutil.WebServiceManager svcmgr)
			{
				// properties can come from files or items.
				foreach(string entityClass in new string[] { "FILE", "ITEM" })
				{
					m_propDefsByEntityClassAndId.Add(
						entityClass,
						svcmgr.PropertyService.GetPropertyDefinitionsByEntityClassId(entityClass).ToDictionary(pd => pd.Id)
						);
				}
			}

			/// <summary>
			/// Sync properties of a file.
			/// </summary>
			/// <param name="svcmgr">a WebServiceManager</param>
			/// <param name="file">the file you would like to sync</param>
			/// <param name="comment">the comment for the new version (if a property sync was performed)</param>
			/// <param name="allowSync">if the local filestore doesn't have the file, get it from another filestore</param>
			/// <param name="writeResults">see FilestoreService.CopyFile method</param>
			/// <param name="cloakedEntityClasses">if you can't read an entity where properties would come from, its entity class is returned here</param>
			/// <param name="force">skip check for equivalence and always do a sync, creating a new version</param>
			/// <returns>the file returned from the checkin, same as the input if no property sync is done</returns>
			public vltobj.File SyncProperties(vltutil.WebServiceManager svcmgr, vltobj.File file, string comment, bool allowSync, out vltobj.PropWriteResults writeResults, out string[] cloakedEntityClasses, bool force=false)
			{
				// NOTE: we could have held onto svcmgr passed to constructor, 
				// but holding onto to something like that is generally a bad idea.

				// clear output parameters so we don't have to worry about that for each possible return condition.
				writeResults = null;
				cloakedEntityClasses = null;

				// first check for property compliance failures.
				// We don't need to sync unless there are equivalence failures.
				if (!force) // skip this fast-out if we are forcing a sync.
				{
					// NOTE: if we synced props to multiple files at a time, we could get compliance failures for all of them in one call.
					vltobj.PropCompFail[] complianceFailures = svcmgr.PropertyService.GetPropertyComplianceFailuresByEntityIds(
						"FILE", new long[] { file.Id }, /*filterPending*/true
						);
					if (complianceFailures == null 
						|| complianceFailures.Sum(cf => (cf.PropEquivFailArray != null ? cf.PropEquivFailArray.Length : 0)) == 0
						)
					{
						// nothing to do!
						return file;
					}
				}

				// checkout file without downloading it.
				// this will fail if file is already checked out.
				vltobj.ByteArray downloadTicket;
				file = svcmgr.DocumentService.CheckoutFile(
					file.Id, vltobj.CheckoutFileOptions.Master,
					/*machine*/Environment.MachineName, /*localPath*/string.Empty, comment,
					out downloadTicket
					);

				try // if anything goes wrong from here on out, undo the checkout
				{
					// get component properties.
					// NOTE: Behavior of this API changed significantly in 2015-R2.
					//       This will not work correctly in 2015 or earlier!
					// NOTE: a null component UID means to get write-back properties for root component in the file.
					// WARNING: we can't sync component level properties without CAD!
					vltobj.CompProp[] compProps = svcmgr.DocumentService.GetComponentProperties(file.Id, /*compUID*/null);

					// return cloaked entity classes thru out parameter.
					// NOTE: a propDefId of -1 indicates get couldn't get properties from an inaccessible entity.
					cloakedEntityClasses = compProps.Where(p => p.PropDefId < 0).Select(p => p.EntClassId).ToArray();
					if (cloakedEntityClasses != null && cloakedEntityClasses.Length > 0)
					{
						// don't proceed since we don't have the permissions to write back 
						// everything that is necessary to clear the failures.
						return svcmgr.DocumentService.UndoCheckoutFile(file.MasterId, out downloadTicket);
					}

					// filter so we only keep values from non-cloaked entities
					// NOTE: this is unnecessary as long as we bail out if there are cloaked entities involved.
					compProps = compProps.Where(p => p.PropDefId > 0).ToArray();
					
					// if there is nothing to write back, bail out.
					// We shouldn't have made it this far if this was the case;
					// but you never can be too sure!
					if (compProps == null || compProps.Length == 0)
					{
						// nothing to do, undo checkout and return
						return svcmgr.DocumentService.UndoCheckoutFile(file.MasterId, out downloadTicket);
					}

					// convert CompProp array to PropWriteReq array.
					//	It really sucks that the Value is a string - it should be an object.
					//	Since it is an invariant string, we need to parse it using invariant locales and convert dates from UTC.
					//	To do that we need to know what the type is, which we can get from the PropertyDef for the PropDefId.
					//	Since we don't want to get property defs everytime this is called, we use a cache.
					//	This should be fixed in a future release.
					vltobj.PropWriteReq[] writeProps = compProps.Select(
						p => new vltobj.PropWriteReq()
						{
							Moniker = p.Moniker,
							CanCreate = p.CreateNew,
							Val = p.Val
						}
						).ToArray();

					// use CopyFile to copy existing resource and write the properties.
					byte[] uploadTicket = svcmgr.FilestoreService.CopyFile(
						downloadTicket.Bytes, allowSync, writeProps,
						out writeResults
						);

					// get child file associations so we can preserve them.
					// NOTE: if we synced props to multiple files at a time, we could get file associations for all of them in one call.
					vltobj.FileAssocLite[] childAssocs = svcmgr.DocumentService.GetFileAssociationLitesByIds(
						new long[] { file.Id },
						vltobj.FileAssocAlg.Actual, // preserve the associations provided by CAD add-in
						/*parentAssociationType*/vltobj.FileAssociationTypeEnum.None, /*parentRecurse*/false,
						/*childAssociationType*/vltobj.FileAssociationTypeEnum.All, /*childRecurse*/false,
						/*includeLibraryFiles*/true,
						/*includeRelatedDocuments*/false,
						/*includeHidden*/true
						);
					// convert FileAssocLite array to FileAssocParam array
					vltobj.FileAssocParam[] associations = childAssocs.Select(
						a => new vltobj.FileAssocParam()
							{
								Typ = a.Typ, CldFileId = a.CldFileId,
								Source = a.Source, RefId = a.RefId, 
								ExpectedVaultPath = a.ExpectedVaultPath
							}
						).ToArray();
					// checkin file
					file = svcmgr.DocumentService.CheckinUploadedFile(
						file.MasterId,
						comment, /*keepCheckedOut*/false, /*lastWrite*/DateTime.Now,
						associations,
						/*bom*/null, /*copyBom*/true, // preserve any BOM
						file.Name, file.FileClass, file.Hidden, // preserve these attributes
						new vltobj.ByteArray() { Bytes = uploadTicket }
						);
				}
				finally
				{
					// if we got here and file is still checked-out, 
					// something went wrong so undo the checkout.
					if (file.CheckedOut)
						file = svcmgr.DocumentService.UndoCheckoutFile(file.MasterId, out downloadTicket);
				}

				return file;
			}

			private static object ConvertInvariantStringToObject(string value, vltobj.DataType typ)
			{
				if (value == null) return null; // quick exit on null

				// dates need to parsed with invariant culture and converted to local time
				if (typ == vltobj.DataType.DateTime)
					return DateTime.Parse(value, System.Globalization.CultureInfo.InvariantCulture).ToLocalTime();
				// numerics need to be parsed with invariant culture
				else if (typ == vltobj.DataType.Numeric)
					return Double.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
				// bools are straight forward
				else if (typ == vltobj.DataType.Bool)
					return Boolean.Parse(value);
				else // don't do anything for string or image types (image types should not be seen here)
					return value;
			}
		}
	}
'@

$2020AndNewerCs = @'
	using System;
	using System.Collections.Generic;
	using System.Linq;

	// Vault SDK
	using vltobj = Autodesk.Connectivity.WebServices;
	using vltutil = Autodesk.Connectivity.WebServicesTools;

	namespace Autodesk.Vault.SDKTools
	{
		/// <summary>
		/// Helper class to do a property sync using filestore service.
		/// This code has been tested with the 2016 Vault API. It should also 
		/// work for 2015-R2 API, but it will not work with 2015 or earlier API.
		/// </summary>
		/// <author>Dave Mink</author>
		public class PropertySync
		{
			// property def cache
			// Needed in order to determine how to parse the invariant string value we get from GetComponentProperties
			private Dictionary<string, Dictionary<long, vltobj.PropDef>> m_propDefsByEntityClassAndId = new Dictionary<string, Dictionary<long, vltobj.PropDef>>();

			public PropertySync(vltutil.WebServiceManager svcmgr)
			{
				// properties can come from files or items.
				foreach(string entityClass in new string[] { "FILE", "ITEM" })
				{
					m_propDefsByEntityClassAndId.Add(
						entityClass,
						svcmgr.PropertyService.GetPropertyDefinitionsByEntityClassId(entityClass).ToDictionary(pd => pd.Id)
						);
				}
			}

			/// <summary>
			/// Sync properties of a file.
			/// </summary>
			/// <param name="svcmgr">a WebServiceManager</param>
			/// <param name="file">the file you would like to sync</param>
			/// <param name="comment">the comment for the new version (if a property sync was performed)</param>
			/// <param name="allowSync">if the local filestore doesn't have the file, get it from another filestore</param>
			/// <param name="writeResults">see FilestoreService.CopyFile method</param>
			/// <param name="cloakedEntityClasses">if you can't read an entity where properties would come from, its entity class is returned here</param>
			/// <param name="force">skip check for equivalence and always do a sync, creating a new version</param>
			/// <returns>the file returned from the checkin, same as the input if no property sync is done</returns>
			public vltobj.File SyncProperties(vltutil.WebServiceManager svcmgr, vltobj.File file, string comment, bool allowSync, out vltobj.PropWriteResults writeResults, out string[] cloakedEntityClasses, bool force=false)
			{
				// NOTE: we could have held onto svcmgr passed to constructor, 
				// but holding onto to something like that is generally a bad idea.

				// clear output parameters so we don't have to worry about that for each possible return condition.
				writeResults = null;
				cloakedEntityClasses = null;

				// first check for property compliance failures.
				// We don't need to sync unless there are equivalence failures.
				if (!force) // skip this fast-out if we are forcing a sync.
				{
					// NOTE: if we synced props to multiple files at a time, we could get compliance failures for all of them in one call.
					vltobj.PropCompFail[] complianceFailures = svcmgr.PropertyService.GetPropertyComplianceFailuresByEntityIds(
						"FILE", new long[] { file.Id }, /*filterPending*/true
						);
					if (complianceFailures == null 
						|| complianceFailures.Sum(cf => (cf.PropEquivFailArray != null ? cf.PropEquivFailArray.Length : 0)) == 0
						)
					{
						// nothing to do!
						return file;
					}
				}

				// checkout file without downloading it.
				// this will fail if file is already checked out.
				vltobj.ByteArray downloadTicket;
				file = svcmgr.DocumentService.CheckoutFile(
					file.Id, vltobj.CheckoutFileOptions.Master,
					/*machine*/Environment.MachineName, /*localPath*/string.Empty, comment,
					out downloadTicket
					);

				try // if anything goes wrong from here on out, undo the checkout
				{
					// get component properties.
					// NOTE: Behavior of this API changed significantly in 2015-R2.
					//       This will not work correctly in 2015 or earlier!
					// NOTE: a null component UID means to get write-back properties for root component in the file.
					// WARNING: we can't sync component level properties without CAD!
					vltobj.CompProp[] compProps = svcmgr.DocumentService.GetComponentProperties(file.Id, /*compUID*/null);

					// return cloaked entity classes thru out parameter.
					// NOTE: a propDefId of -1 indicates get couldn't get properties from an inaccessible entity.
					cloakedEntityClasses = compProps.Where(p => p.PropDefId < 0).Select(p => p.EntClassId).ToArray();
					if (cloakedEntityClasses != null && cloakedEntityClasses.Length > 0)
					{
						// don't proceed since we don't have the permissions to write back 
						// everything that is necessary to clear the failures.
						return svcmgr.DocumentService.UndoCheckoutFile(file.MasterId, out downloadTicket);
					}

					// filter so we only keep values from non-cloaked entities
					// NOTE: this is unnecessary as long as we bail out if there are cloaked entities involved.
					compProps = compProps.Where(p => p.PropDefId > 0).ToArray();
					
					// if there is nothing to write back, bail out.
					// We shouldn't have made it this far if this was the case;
					// but you never can be too sure!
					if (compProps == null || compProps.Length == 0)
					{
						// nothing to do, undo checkout and return
						return svcmgr.DocumentService.UndoCheckoutFile(file.MasterId, out downloadTicket);
					}

					// convert CompProp array to PropWriteReq array.
					//	It really sucks that the Value is a string - it should be an object.
					//	Since it is an invariant string, we need to parse it using invariant locales and convert dates from UTC.
					//	To do that we need to know what the type is, which we can get from the PropertyDef for the PropDefId.
					//	Since we don't want to get property defs everytime this is called, we use a cache.
					//	This should be fixed in a future release.
					vltobj.PropWriteReq[] writeProps = compProps.Select(
						p => new vltobj.PropWriteReq()
						{
							Moniker = p.Moniker,
							CanCreate = p.CreateNew,
							Val = p.Val
						}
						).ToArray();

					// use CopyFile to copy existing resource and write the properties.
					byte[] uploadTicket = svcmgr.FilestoreService.CopyFile(
						downloadTicket.Bytes, null, allowSync, writeProps,
						out writeResults
						);

					// get child file associations so we can preserve them.
					// NOTE: if we synced props to multiple files at a time, we could get file associations for all of them in one call.
					vltobj.FileAssocLite[] childAssocs = svcmgr.DocumentService.GetFileAssociationLitesByIds(
						new long[] { file.Id },
						vltobj.FileAssocAlg.Actual, // preserve the associations provided by CAD add-in
						/*parentAssociationType*/vltobj.FileAssociationTypeEnum.None, /*parentRecurse*/false,
						/*childAssociationType*/vltobj.FileAssociationTypeEnum.All, /*childRecurse*/false,
						/*includeLibraryFiles*/true,
						/*includeRelatedDocuments*/false,
						/*includeHidden*/true
						);
					// convert FileAssocLite array to FileAssocParam array
					vltobj.FileAssocParam[] associations = childAssocs.Select(
						a => new vltobj.FileAssocParam()
							{
								Typ = a.Typ, CldFileId = a.CldFileId,
								Source = a.Source, RefId = a.RefId, 
								ExpectedVaultPath = a.ExpectedVaultPath
							}
						).ToArray();
					// checkin file
					file = svcmgr.DocumentService.CheckinUploadedFile(
						file.MasterId,
						comment, /*keepCheckedOut*/false, /*lastWrite*/DateTime.Now,
						associations,
						/*bom*/null, /*copyBom*/true, // preserve any BOM
						file.Name, file.FileClass, file.Hidden, // preserve these attributes
						new vltobj.ByteArray() { Bytes = uploadTicket }
						);
				}
				finally
				{
					// if we got here and file is still checked-out, 
					// something went wrong so undo the checkout.
					if (file.CheckedOut)
						file = svcmgr.DocumentService.UndoCheckoutFile(file.MasterId, out downloadTicket);
				}

				return file;
			}

			private static object ConvertInvariantStringToObject(string value, vltobj.DataType typ)
			{
				if (value == null) return null; // quick exit on null

				// dates need to parsed with invariant culture and converted to local time
				if (typ == vltobj.DataType.DateTime)
					return DateTime.Parse(value, System.Globalization.CultureInfo.InvariantCulture).ToLocalTime();
				// numerics need to be parsed with invariant culture
				else if (typ == vltobj.DataType.Numeric)
					return Double.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
				// bools are straight forward
				else if (typ == vltobj.DataType.Bool)
					return Boolean.Parse(value);
				else // don't do anything for string or image types (image types should not be seen here)
					return value;
			}
		}
	}
'@

$2022AndNewerCs = @'
	using System;
	using System.Collections.Generic;
	using System.Linq;

	// Vault SDK
	using vltobj = Autodesk.Connectivity.WebServices;
	using vltutil = Autodesk.Connectivity.WebServicesTools;

	namespace Autodesk.Vault.SDKTools
	{
		/// <summary>
		/// Helper class to do a property sync using filestore service.
		/// This code has been tested with the 2016 Vault API. It should also 
		/// work for 2015-R2 API, but it will not work with 2015 or earlier API.
		/// </summary>
		/// <author>Dave Mink</author>
		public class PropertySync
		{
			// property def cache
			// Needed in order to determine how to parse the invariant string value we get from GetComponentProperties
			private Dictionary<string, Dictionary<long, vltobj.PropDef>> m_propDefsByEntityClassAndId = new Dictionary<string, Dictionary<long, vltobj.PropDef>>();

			public PropertySync(vltutil.WebServiceManager svcmgr)
			{
				// properties can come from files or items.
				foreach(string entityClass in new string[] { "FILE", "ITEM" })
				{
					m_propDefsByEntityClassAndId.Add(
						entityClass,
						svcmgr.PropertyService.GetPropertyDefinitionsByEntityClassId(entityClass).ToDictionary(pd => pd.Id)
						);
				}
			}

			/// <summary>
			/// Sync properties of a file.
			/// </summary>
			/// <param name="svcmgr">a WebServiceManager</param>
			/// <param name="file">the file you would like to sync</param>
			/// <param name="comment">the comment for the new version (if a property sync was performed)</param>
			/// <param name="allowSync">if the local filestore doesn't have the file, get it from another filestore</param>
			/// <param name="writeResults">see FilestoreService.CopyFile method</param>
			/// <param name="cloakedEntityClasses">if you can't read an entity where properties would come from, its entity class is returned here</param>
			/// <param name="force">skip check for equivalence and always do a sync, creating a new version</param>
			/// <returns>the file returned from the checkin, same as the input if no property sync is done</returns>
			public vltobj.File SyncProperties(vltutil.WebServiceManager svcmgr, vltobj.File file, string comment, bool allowSync, out vltobj.PropWriteResults writeResults, out string[] cloakedEntityClasses, bool force=false)
			{
				// NOTE: we could have held onto svcmgr passed to constructor, 
				// but holding onto to something like that is generally a bad idea.

				// clear output parameters so we don't have to worry about that for each possible return condition.
				writeResults = null;
				cloakedEntityClasses = null;

				// first check for property compliance failures.
				// We don't need to sync unless there are equivalence failures.
				if (!force) // skip this fast-out if we are forcing a sync.
				{
					// NOTE: if we synced props to multiple files at a time, we could get compliance failures for all of them in one call.
					vltobj.PropCompFail[] complianceFailures = svcmgr.PropertyService.GetPropertyComplianceFailuresByEntityIds(
						"FILE", new long[] { file.Id }, /*filterPending*/true
						);
					if (complianceFailures == null 
						|| complianceFailures.Sum(cf => (cf.PropEquivFailArray != null ? cf.PropEquivFailArray.Length : 0)) == 0
						)
					{
						// nothing to do!
						return file;
					}
				}

				// checkout file without downloading it.
				// this will fail if file is already checked out.
				vltobj.ByteArray downloadTicket;
				file = svcmgr.DocumentService.CheckoutFile(
					file.Id, vltobj.CheckoutFileOptions.Master,
					/*machine*/Environment.MachineName, /*localPath*/string.Empty, comment,
					out downloadTicket
					);

				try // if anything goes wrong from here on out, undo the checkout
				{
					// get component properties.
					// NOTE: Behavior of this API changed significantly in 2015-R2.
					//       This will not work correctly in 2015 or earlier!
					// NOTE: a null component UID means to get write-back properties for root component in the file.
					// WARNING: we can't sync component level properties without CAD!
					vltobj.CompProp[] compProps = svcmgr.DocumentService.GetComponentProperties(file.Id, /*compUID*/null);

					// return cloaked entity classes thru out parameter.
					// NOTE: a propDefId of -1 indicates get couldn't get properties from an inaccessible entity.
					cloakedEntityClasses = compProps.Where(p => p.PropDefId < 0).Select(p => p.EntClassId).ToArray();
					if (cloakedEntityClasses != null && cloakedEntityClasses.Length > 0)
					{
						// don't proceed since we don't have the permissions to write back 
						// everything that is necessary to clear the failures.
						return svcmgr.DocumentService.UndoCheckoutFile(file.MasterId, out downloadTicket);
					}

					// filter so we only keep values from non-cloaked entities
					// NOTE: this is unnecessary as long as we bail out if there are cloaked entities involved.
					compProps = compProps.Where(p => p.PropDefId > 0).ToArray();
					
					// if there is nothing to write back, bail out.
					// We shouldn't have made it this far if this was the case;
					// but you never can be too sure!
					if (compProps == null || compProps.Length == 0)
					{
						// nothing to do, undo checkout and return
						return svcmgr.DocumentService.UndoCheckoutFile(file.MasterId, out downloadTicket);
					}

					// convert CompProp array to PropWriteReq array.
					//	It really sucks that the Value is a string - it should be an object.
					//	Since it is an invariant string, we need to parse it using invariant locales and convert dates from UTC.
					//	To do that we need to know what the type is, which we can get from the PropertyDef for the PropDefId.
					//	Since we don't want to get property defs everytime this is called, we use a cache.
					//	This should be fixed in a future release.
					vltobj.PropWriteReq[] writeProps = compProps.Select(
						p => new vltobj.PropWriteReq()
						{
							Moniker = p.Moniker,
							CanCreate = p.CreateNew,
							Val = p.Val
						}
						).ToArray();

					vltobj.PropWriteRequests writePropsReq = new vltobj.PropWriteRequests();
					writePropsReq.Requests = writeProps;
					writePropsReq.Bom = null;
					// use CopyFile to copy existing resource and write the properties.
					byte[] uploadTicket = svcmgr.FilestoreService.CopyFile(
						downloadTicket.Bytes, null, allowSync, writePropsReq,
						out writeResults
						);

					// get child file associations so we can preserve them.
					// NOTE: if we synced props to multiple files at a time, we could get file associations for all of them in one call.
					vltobj.FileAssocLite[] childAssocs = svcmgr.DocumentService.GetFileAssociationLitesByIds(
						new long[] { file.Id },
						vltobj.FileAssocAlg.Actual, // preserve the associations provided by CAD add-in
						/*parentAssociationType*/vltobj.FileAssociationTypeEnum.None, /*parentRecurse*/false,
						/*childAssociationType*/vltobj.FileAssociationTypeEnum.All, /*childRecurse*/false,
						/*includeLibraryFiles*/true,
						/*includeRelatedDocuments*/false,
						/*includeHidden*/true
						);
					// convert FileAssocLite array to FileAssocParam array
					vltobj.FileAssocParam[] associations = childAssocs.Select(
						a => new vltobj.FileAssocParam()
							{
								Typ = a.Typ, CldFileId = a.CldFileId,
								Source = a.Source, RefId = a.RefId, 
								ExpectedVaultPath = a.ExpectedVaultPath
							}
						).ToArray();
					// checkin file
					file = svcmgr.DocumentService.CheckinUploadedFile(
						file.MasterId,
						comment, /*keepCheckedOut*/false, /*lastWrite*/DateTime.Now,
						associations,
						/*bom*/null, /*copyBom*/true, // preserve any BOM
						file.Name, file.FileClass, file.Hidden, // preserve these attributes
						new vltobj.ByteArray() { Bytes = uploadTicket }
						);
				}
				finally
				{
					// if we got here and file is still checked-out, 
					// something went wrong so undo the checkout.
					if (file.CheckedOut)
						file = svcmgr.DocumentService.UndoCheckoutFile(file.MasterId, out downloadTicket);
				}

				return file;
			}

			private static object ConvertInvariantStringToObject(string value, vltobj.DataType typ)
			{
				if (value == null) return null; // quick exit on null

				// dates need to parsed with invariant culture and converted to local time
				if (typ == vltobj.DataType.DateTime)
					return DateTime.Parse(value, System.Globalization.CultureInfo.InvariantCulture).ToLocalTime();
				// numerics need to be parsed with invariant culture
				else if (typ == vltobj.DataType.Numeric)
					return Double.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
				// bools are straight forward
				else if (typ == vltobj.DataType.Bool)
					return Boolean.Parse(value);
				else // don't do anything for string or image types (image types should not be seen here)
					return value;
			}
		}
	}
'@
#endregion

function SyncProperties
{
    param($file)

	Write-Host "Start 'Sync Properties' of file '$($file.Name)' ..."
	$s = "/AutodeskDM/Services/v"
	$start = $vault.AdminService.Url.IndexOf($s) + $s.Length
	$version = [int]$vault.AdminService.Url.Substring($start, 2)

	# 27 = 2022
	# 25 = 2020
	# 24 = 2019
	# 23 = 2018
	# 22 = 2017
	# 21 = 2016
	
	if ($version -ge 27) {
		$cs = $2022AndNewerCs
	} elseif ($version -ge 25) {
		$cs = $2020AndNewerCs
	} elseif ($version -ge 22) {
		$cs = $2017AndNewerCs
	} else {
		$cs = $OlderThan2017Cs
	}

	$null = [System.Reflection.Assembly]::LoadWithPartialName("Autodesk.Connectivity.WebServices")
	Add-Type -TypeDefinition $cs -ReferencedAssemblies @("Autodesk.Connectivity.WebServices", "System.Core")
	
	$propertySync = New-Object Autodesk.Vault.SDKTools.PropertySync -ArgumentList @($vault)
	$writeResults = New-Object Autodesk.Connectivity.WebServices.PropWriteResults
	$cloakedEntityClasses = @()
	$vaultFile = $vault.DocumentService.GetFileById($file.Id)
	$vaultFile = $propertySync.SyncProperties($vault, $vaultFile, "Property Sync", $true, [ref]$writeResults, [ref]$cloakedEntityClasses, $true)
	Write-Host "End 'Sync Properties' of file '$($file.Name)'"
}
function UpdateRevisionBlock {
	param ($file)

	# UpdateRevisionBlock needs the latest version of the file
	$fileForUpdRevJob = $vault.DocumentService.GetLatestFileByMasterId($file.MasterId)

	Write-Host "Start 'Update Revision Block' of file '$($fileForUpdRevJob.Name)' ..."

	$ConnectivityJobProcessorDelegateAssembly = [System.Reflection.Assembly]::Load("Connectivity.JobProcessor.Delegate")
	$context = $ConnectivityJobProcessorDelegateAssembly.CreateInstance("Connectivity.JobHandlers.Services.Objects.ServiceJobProcessorServices",$true, [System.Reflection.BindingFlags]::CreateInstance, $null, $null, $null, $null)
	$context.GetType().GetProperty("Connection").SetValue($context, $vaultConnection, $null) #need to set the Connection property on the context or else it runs into error

	$JobHandlerURBAssembly = [System.Reflection.Assembly]::Load("Connectivity.Explorer.JobHandlerUpdateRevisionBlock")
	$uRBJobHandler = $JobHandlerURBAssembly.CreateInstance("Connectivity.Explorer.JobHandlerUpdateRevisionBlock.UpdateRevisionBlockJobHandler",$true, [System.Reflection.BindingFlags]::CreateInstance, $null, $null, $null, $null)
	$uRBJob = New-Object Connectivity.Services.Job.UpdateRevisionBlockJob("vault",$fileForUpdRevJob.Id,$false,$false,$fileForUpdRevJob.Name)

	$jobOutcome = $uRBJobHandler.Execute($context,$uRBJob) #call execute to start running the job

	if ($jobOutcome -eq "Failure")
	{
		throw "Failed job 'Update Revision Block'" #Failed because of issue that occured in the job
	}
	Write-Host "End 'Update Revision Block' of file '$($fileForUpdRevJob.Name)'"


# Error using this code block: Exception calling "GetFileVersionsByIDs" with "2" argument(s): "1013"

	# $versionIds = @($fileForUpdRevJob.Id)
	# $fileVersions = [Connectivity.Services.Document.DocServices]::Instance.GetFileVersionsByIDs($vaultConnection, $versionIds)
	# $multiSideProvider = new-object Connectivity.Explorer.JobHandlerUpdateRevisionBlock.URBJobHandlerMultiSiteSyncProvider -ArgumentList $true
	# $updateRevisionBlock = new-object Connectivity.Explorer.Document.ViewModel.UpdateRevisionBlock -ArgumentList $fileVersions, $false
	# $updateRevisionBlock.MultiSiteSyncProvider = $multiSideProvider
	# $updateRevisionBlock.CheckInComment = "Revision table updated by coolOrange UpdateRevisionTable"
	# $updateRevisionBlock.RefreshRevisionBlock()

}

