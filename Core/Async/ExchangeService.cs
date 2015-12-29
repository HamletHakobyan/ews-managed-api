/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.Generic;
    using System.Threading.Tasks;

    /// <summary>
    /// Represents a binding to the Exchange Web Services.
    /// </summary>
    public sealed partial class ExchangeService
    {
        /// <summary>
        /// Binds to multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <param name="anchorMailbox">The SmtpAddress of mailbox that hosts all items we need to bind to</param>
        /// <param name="errorHandling">Type of error handling to perform.</param>
        /// <returns>A ServiceResponseCollection providing results for each of the specified item Ids.</returns>
        private Task<ServiceResponseCollection<GetItemResponse>> InternalBindToItemsAsync(
            IEnumerable<ItemId> itemIds,
            PropertySet propertySet,
            string anchorMailbox,
            ServiceErrorHandling errorHandling)
        {
            GetItemRequest request = new GetItemRequest(this, errorHandling);

            request.ItemIds.AddRange(itemIds);
            request.PropertySet = propertySet;
            request.AnchorMailbox = anchorMailbox;

            return request.ExecuteAsync();
        }

        /// <summary>
        /// Binds to item.
        /// </summary>
        /// <param name="itemId">The item id.</param>
        /// <param name="propertySet">The property set.</param>
        /// <returns>Item.</returns>
        internal async Task<Item> BindToItemAsync(ItemId itemId, PropertySet propertySet)
        {
            EwsUtilities.ValidateParam(itemId, "itemId");
            EwsUtilities.ValidateParam(propertySet, "propertySet");

            ServiceResponseCollection<GetItemResponse> responses = await this.InternalBindToItemsAsync(
                new ItemId[] { itemId },
                propertySet,
                null, /* anchorMailbox */
                ServiceErrorHandling.ThrowOnError);

            return responses[0].Item;
        }

        /// <summary>
        /// Binds to item.
        /// </summary>
        /// <typeparam name="TItem">The type of the item.</typeparam>
        /// <param name="itemId">The item id.</param>
        /// <param name="propertySet">The property set.</param>
        /// <returns>Item</returns>
        internal async Task<TItem> BindToItemAsync<TItem>(ItemId itemId, PropertySet propertySet)
            where TItem : Item
        {
            Item result = await this.BindToItemAsync(itemId, propertySet);

            if (result is TItem)
            {
                return (TItem)result;
            }
            else
            {
                throw new ServiceLocalException(
                    string.Format(
                        Strings.ItemTypeNotCompatible,
                        result.GetType().Name,
                        typeof(TItem).Name));
            }
        }
    }
}
