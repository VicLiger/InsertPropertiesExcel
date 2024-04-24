using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Navisworks.Api;
using ComApiBridge = Autodesk.Navisworks.Api.Interop.ComApi;
using ComApiBridge2 = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
using Autodesk.Navisworks.Api.Interop.ComApi;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
using System.Windows.Forms;

namespace Properties.Methods
{
    static class PropertiesInsert
    {
        private static ComApi.InwOpState10 _oState = ComApiBridge2.State;
        public static void PropertyInsertion(string catName, string propName, string propValue, ModelItem selModel)
        {
            int _index = 1;
            ComApi.InwOaPath convModel = ComApiBridge2.ToInwOaPath(selModel);
            ComApiBridge.InwGUIPropertyNode2 propCat = (ComApiBridge.InwGUIPropertyNode2)_oState.GetGUIPropertyNode(convModel, true);
            ComApiBridge.InwGUIAttributesColl propCol = propCat.GUIAttributes();
            foreach (ComApiBridge.InwGUIAttribute2 attr in propCol)
            {
                if (attr.UserDefined)
                {
                    if (attr.ClassUserName == catName)
                    {
                        try
                        {
                            propCat.SetUserDefined(_index, catName, $"{catName}_InternalName", AddNewPropertyToExtgCategory(attr, propName, propValue));
                            return;
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show("Houve um problema ao tentar criar a categoria");
                            return;
                        }
                    }
                    _index++;
                }
            }
            FirstPropertyInsertion(catName, propName, propValue, new ModelItemCollection() { selModel }, 0);
        }
        private static ComApiBridge.InwOaPropertyVec AddNewPropertyToExtgCategory(ComApiBridge.InwGUIAttribute2 propertCat, string PropertyName, string PropertyValue)
        {
            ComApiBridge.InwOpState10 cdoc = ComApiBridge2.State;
            ComApiBridge.InwOaPropertyVec category = (ComApiBridge.InwOaPropertyVec)cdoc.ObjectFactory(ComApiBridge.nwEObjectType.eObjectType_nwOaPropertyVec, null, null);
            foreach (ComApiBridge.InwOaProperty property in propertCat.Properties())
            {
                ComApiBridge.InwOaProperty extgProp = (ComApiBridge.InwOaProperty)cdoc.ObjectFactory(ComApiBridge.nwEObjectType.eObjectType_nwOaProperty, null, null);
                extgProp.name = property.name;
                extgProp.UserName = property.UserName;
                extgProp.value = property.value;
                if (extgProp.UserName != PropertyName) category.Properties().Add(extgProp);
            }
            ComApiBridge.InwOaProperty newProp = (ComApiBridge.InwOaProperty)cdoc.ObjectFactory(ComApiBridge.nwEObjectType.eObjectType_nwOaProperty, null, null);
            newProp.name = $"{PropertyName}_Internal";
            newProp.UserName = PropertyName;
            newProp.value = PropertyValue;
            category.Properties().Add(newProp);
            return category;
        }
        private static void FirstPropertyInsertion(string catName, string propName, string propValue, ModelItemCollection selectionModel, int index)
        {
            try
            {
                if (selectionModel.Count > 0)
                {
                    ComApi.InwOpSelection comSelectionOut = ComApiBridge2.ToInwOpSelection(selectionModel);
                    //create a new propety and add it to the category
                    ComApi.InwOaProperty customProp = (ComApi.InwOaProperty)_oState.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaProperty, null, null);
                    customProp.name = $"{catName}_Internal";
                    customProp.UserName = propName;
                    customProp.value = propValue;
                    #region Creating a New PropVector
                    ComApi.InwOaPropertyVec _propVector = (ComApi.InwOaPropertyVec)_oState.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaPropertyVec, null, null);
                    #endregion
                    _propVector.Properties().Add(customProp);
                    foreach (ComApi.InwOaPath3 oPath in comSelectionOut.Paths())
                    {
                        ComApi.InwOaPath3 oPath1 = oPath;
                        ComApi.InwGUIPropertyNode2 propn = (ComApi.InwGUIPropertyNode2)_oState.GetGUIPropertyNode(oPath1, true);
                        //add the new category to the object
                        propn.SetUserDefined(index, catName, $"{catName}_Internal", _propVector);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocorreu um erro: " + ex.Message);
            }
        }
    }
}
